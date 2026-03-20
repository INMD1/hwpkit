/**
 * XmlAdapter - fast-xml-parser 기반 XML 처리 서비스
 *
 * 기존 브라우저 네이티브 DOMParser 대신 fast-xml-parser를 사용합니다.
 * JS 객체 트리로 파싱하므로 DOM 노드 순회 방식과 데이터 흐름이 완전히 다릅니다.
 */

import { XMLParser, XMLBuilder } from 'fast-xml-parser';

// ─── Parser ───────────────────────────────────────────────────────────────────

/** fast-xml-parser가 파싱한 결과 타입 */
export type XmlObject = Record<string, unknown>;

export interface XmlParseOptions {
  /** 배열로 처리할 태그명 목록 */
  arrayTags?: string[];
  /** 속성 접두어 (기본값: '@_') */
  attrPrefix?: string;
  /** 네임스페이스 접두어 제거 여부 */
  removeNamespace?: boolean;
}

export function parseXml(xmlText: string, opts: XmlParseOptions = {}): XmlObject {
  const { arrayTags = [], attrPrefix = '@_', removeNamespace = false } = opts;
  const tagSet = new Set(arrayTags);

  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: attrPrefix,
    textNodeName: '#text',
    removeNSPrefix: removeNamespace,
    isArray: (tagName) => tagSet.has(tagName),
    parseAttributeValue: false,
    trimValues: false,
    cdataPropName: '#cdata',
  });

  try {
    return parser.parse(xmlText) as XmlObject;
  } catch (e) {
    throw new Error(`[XmlAdapter] XML 파싱 실패: ${(e as Error).message}`);
  }
}

// ─── Value Accessors ──────────────────────────────────────────────────────────

/** 중첩 경로로 값 접근 (예: 'hh:sec.hp:p') */
export function getPath(obj: XmlObject, path: string): unknown {
  return path.split('.').reduce<unknown>((cur, key) => {
    if (cur == null || typeof cur !== 'object') return undefined;
    return (cur as Record<string, unknown>)[key];
  }, obj);
}

export function asArray<T>(value: unknown): T[] {
  if (value == null) return [];
  if (Array.isArray(value)) return value as T[];
  return [value as T];
}

export function asString(value: unknown, fallback = ''): string {
  if (value == null) return fallback;
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (typeof value === 'object' && '#text' in (value as object)) {
    return String((value as Record<string, unknown>)['#text'] ?? fallback);
  }
  return fallback;
}

export function asNumber(value: unknown, fallback = 0): number {
  const n = Number(value);
  return isNaN(n) ? fallback : n;
}

export function attr(obj: XmlObject, name: string, prefix = '@_'): string {
  return asString((obj as Record<string, unknown>)[`${prefix}${name}`]);
}

export function numAttr(obj: XmlObject, name: string, prefix = '@_'): number {
  return asNumber((obj as Record<string, unknown>)[`${prefix}${name}`]);
}

// ─── Builder ──────────────────────────────────────────────────────────────────

export interface XmlBuildOptions {
  attrPrefix?: string;
  declaration?: boolean;
}

export function buildXml(obj: XmlObject, opts: XmlBuildOptions = {}): string {
  const { attrPrefix = '@_', declaration = true } = opts;
  const builder = new XMLBuilder({
    ignoreAttributes: false,
    attributeNamePrefix: attrPrefix,
    textNodeName: '#text',
    format: false,
    suppressBooleanAttributes: false,
  });
  const xml = builder.build(obj) as string;
  return declaration
    ? `<?xml version="1.0" encoding="UTF-8"?>\n${xml}`
    : xml;
}

// ─── Simple Tag Builder ───────────────────────────────────────────────────────

/** XML 태그를 문자열로 빠르게 조합하는 유틸 */
export class TagBuilder {
  private parts: string[] = [];

  open(tag: string, attrs: Record<string, string | number | boolean> = {}): this {
    const attrStr = Object.entries(attrs)
      .filter(([, v]) => v !== undefined && v !== '')
      .map(([k, v]) => ` ${k}="${String(v).replace(/"/g, '&quot;')}"`)
      .join('');
    this.parts.push(`<${tag}${attrStr}>`);
    return this;
  }

  close(tag: string): this {
    this.parts.push(`</${tag}>`);
    return this;
  }

  selfClose(tag: string, attrs: Record<string, string | number | boolean> = {}): this {
    const attrStr = Object.entries(attrs)
      .filter(([, v]) => v !== undefined && v !== '')
      .map(([k, v]) => ` ${k}="${String(v).replace(/"/g, '&quot;')}"`)
      .join('');
    this.parts.push(`<${tag}${attrStr}/>`);
    return this;
  }

  text(value: string): this {
    this.parts.push(escapeXml(value));
    return this;
  }

  raw(xml: string): this {
    this.parts.push(xml);
    return this;
  }

  wrap(tag: string, attrs: Record<string, string | number | boolean>, inner: string): this {
    this.open(tag, attrs).raw(inner).close(tag);
    return this;
  }

  build(): string {
    return this.parts.join('');
  }
}

export function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
