/**
 * HWPX 인코더 유틸리티 함수
 *
 * HWPX 인코딩에 필요한 재사용 가능한 함수들을 모아둔 모듈입니다.
 */

import {
  HWPUNIT_PER_PT,
  PT_PER_PIXEL,
  PT_PER_INCH,
} from './constants';

// ==================== 단위 변환 함수 ====================

/**
 * 포인트 (pt) 를 HWPUNIT 으로 변환합니다.
 * @param pt - 포인트 값
 * @returns HWPUNIT 값 (정수)
 * @example
 * ptToHwp(10)  // 10000 (10pt × 1000)
 */
export function ptToHwp(pt: number): number {
  return Math.round(pt * HWPUNIT_PER_PT);
}

/**
 * HWPUNIT 을 포인트 (pt) 로 변환합니다.
 * @param hwpunit - HWPUNIT 값
 * @returns 포인트 값
 * @example
 * hwpToPt(10000)  // 10 (10000 ÷ 1000)
 */
export function hwpToPt(hwpunit: number): number {
  return hwpunit / HWPUNIT_PER_PT;
}

/**
 * 포인트 (pt) 를 HWP 높이 단위로 변환합니다.
 * 폰트 높이로 사용되며 ptToHwp 와 동일합니다.
 * @param pt - 포인트 값
 * @returns HWP 높이 단위
 */
export function ptToHHeight(pt: number): number {
  return ptToHwp(pt);
}

/**
 * 픽셀을 포인트 (pt) 로 변환합니다.
 * @param pixels - 픽셀 값
 * @returns 포인트 값
 * @example
 * pixelsToPt(96)  // 72 (96px × 72/96)
 */
export function pixelsToPt(pixels: number): number {
  return pixels * PT_PER_PIXEL;
}

/**
 * 픽셀을 HWPUNIT 으로 변환합니다.
 * @param pixels - 픽셀 값
 * @returns HWPUNIT 값
 */
export function pixelsToHwp(pixels: number): number {
  return ptToHwp(pixelsToPt(pixels));
}

/**
 * 인치 (inch) 를 포인트 (pt) 로 변환합니다.
 * @param inches - 인치 값
 * @returns 포인트 값
 * @example
 * inchesToPt(1)  // 72 (1 인치 × 72)
 */
export function inchesToPt(inches: number): number {
  return inches * PT_PER_INCH;
}

/**
 * 인치 (inch) 를 HWPUNIT 으로 변환합니다.
 * @param inches - 인치 값
 * @returns HWPUNIT 값
 */
export function inchesToHwp(inches: number): number {
  return ptToHwp(inchesToPt(inches));
}

// ==================== XML 도우미 함수 ====================

/**
 * XML 문자열을 이스케이프 처리합니다.
 * @param str - 이스케이프할 문자열
 * @returns 이스케이프된 문자열
 * @example
 * escapeXml("<hello>&world")  // "&lt;hello&gt;&amp;world"
 */
export function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * 속성 문자열을 생성합니다.
 * @param attributes - 속성 객체 (null/undefined 값은 제외됨)
 * @returns 속성 문자열
 * @example
 * buildAttributes({ id: "1", name: "test", disabled: null })
 * // 'id="1" name="test"'
 */
export function buildAttributes(attributes: Record<string, string | number | null | undefined>): string {
  const parts: string[] = [];
  for (const [key, value] of Object.entries(attributes)) {
    if (value != null && value !== "") {
      const escapedValue = typeof value === "string" ? escapeXml(value) : String(value);
      parts.push(`${key}="${escapedValue}"`);
    }
  }
  return parts.join(" ");
}

/**
 * XML 요소를 생성합니다.
 * @param tagName - 태그 이름
 * @param attributes - 속성 객체
 * @param content - 내용 (문자열 또는 자식 요소 배열)
 * @returns XML 요소 문자열
 * @example
 * createElement("span", { id: "1" }, "내용")
 * // '<span id="1">내용</span>'
 */
export function createElement(
  tagName: string,
  attributes: Record<string, string | number | null | undefined> = {},
  content: string | string[] = "",
): string {
  const attrs = buildAttributes(attributes);
  const attrsStr = attrs ? ` ${attrs}` : "";
  const contentStr = Array.isArray(content) ? content.join("") : content;

  if (contentStr === "") {
    return `<${tagName}${attrsStr}/>`;
  }
  return `<${tagName}${attrsStr}>${contentStr}</${tagName}>`;
}

// ==================== 문자열 처리 함수 ====================

/**
 * 문자열이 한글을 포함하는지 확인합니다.
 * @param text - 확인할 문자열
 * @returns 한글이 포함되어 있으면 true
 */
export function containsKorean(text: string): boolean {
  return /[\uAC00-\uD7A3\u3131-\u318E]/.test(text);
}

/**
 * 문자열이 영문을 포함하는지 확인합니다.
 * @param text - 확인할 문자열
 * @returns 영문이 포함되어 있으면 true
 */
export function containsLatin(text: string): boolean {
  return /[a-zA-Z]/.test(text);
}

// ==================== 계산 함수 ====================

/**
 * 두 숫자 중 큰 값을 반환합니다.
 * @param a - 첫 번째 숫자
 * @param b - 두 번째 숫자
 * @returns 큰 값
 */
export function max(a: number, b: number): number {
  return a > b ? a : b;
}

/**
 * 두 숫자 중 작은 값을 반환합니다.
 * @param a - 첫 번째 숫자
 * @param b - 두 번째 숫자
 * @returns 작은 값
 */
export function min(a: number, b: number): number {
  return a < b ? a : b;
}

/**
 * 숫자를 정수로 반올림합니다.
 * @param value - 변환할 숫자
 * @returns 정수
 */
export function round(value: number): number {
  return Math.round(value);
}
