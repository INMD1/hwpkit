/**
 * 스타일 매핑 유틸리티
 *
 * 다양한 문서 포맷 간의 스타일 (폰트, 색상, 정렬, 테두리 등) 을 매핑합니다.
 */

import type { Align, StrokeKind, Stroke, TextProps, ParaProps } from '../model/doc-props';

// ─── 폰트 매핑 ───────────────────────────────────────────

/** 한글 폰트 ↔ 영문 폰트 매핑 */
export const FONT_MAP: Record<string, string> = {
  // 한글 → 영문
  '맑은 고딕': 'Malgun Gothic',
  '바탕': 'Batang',
  '돋움': 'Dotum',
  '굴림': 'Gulim',
  '한컴바탕': 'Batang',
  '한컴돋움': 'Malgun Gothic',
  '함초롬바탕': 'Batang',
  '함초롬돋움': 'Malgun Gothic',

  // 영문 → 한글 (역매핑)
  'Malgun Gothic': '맑은 고딕',
  'Batang': '바탕',
  'Dotum': '돋움',
  'Gulim': '굴림',
};

/** 폰트를 한글로 변환 (한컴 호환성) */
export function fontToKorean(font?: string): string {
  if (!font) return '맑은 고딕';
  return FONT_MAP[font] ?? font;
}

/** 폰트를 영문으로 변환 (DOCX 호환성) */
export function fontToEnglish(font?: string): string {
  if (!font) return 'Malgun Gothic';

  // 이미 영문 폰트라면 그대로 반환
  if (/^[a-zA-Z\s]+$/.test(font)) {
    return FONT_MAP[font] ?? font;
  }
  // 한글 폰트라면 영문으로 변환
  return FONT_MAP[font] ?? 'Malgun Gothic';
}

// ─── 색상 매핑 ───────────────────────────────────────────

/** 색상을 안전한 6 자리 hex 문자열로 변환 */
export function safeHex(raw: string | number | null | undefined): string | undefined {
  if (raw == null) return undefined;

  if (typeof raw === 'number') {
    if (raw <= 0) return '000000';
    if (raw >= 0xFFFFFF) return undefined;
    return raw.toString(16).padStart(6, '0').toUpperCase();
  }

  let s = String(raw).replace(/^#/, '').toUpperCase();

  // 3 자리 hex → 6 자리 확장
  if (/^[0-9A-F]{3}$/.test(s)) {
    s = s[0] + s[0] + s[1] + s[1] + s[2] + s[2];
  }

  // 6 자리 hex 검증
  if (/^[0-9A-F]{6}$/.test(s)) {
    return s;
  }

  // 특수값 처리
  if (s === 'AUTO' || s === 'NONE' || s === 'TRANSPARENT') {
    return undefined;
  }

  return undefined;
}

// ─── 정렬 매핑 ───────────────────────────────────────────

/** 정렬 값 정규화 */
const ALIGN_MAP: Record<string, Align> = {
  // 대문자
  LEFT: 'left',
  CENTER: 'center',
  RIGHT: 'right',
  JUSTIFY: 'justify',
  BOTH: 'justify',
  DISTRIBUTE: 'justify',

  // 소문자
  left: 'left',
  center: 'center',
  right: 'right',
  both: 'justify',

  // CSS 스타일
  start: 'left',
  end: 'right',
};

export function safeAlign(raw?: string): Align {
  return ALIGN_MAP[raw ?? ''] ?? 'left';
}

// ─── 테두리 매핑 ────────────────────────────────────────

/** HWPX stroke kind 매핑 */
const HWPX_STROKE: Record<string, StrokeKind> = {
  SOLID: 'solid',
  NONE: 'none',
  DASH: 'dash',
  DOT: 'dot',
  DOUBLE: 'double',
  LONG_DASH: 'dash',
  DASH_DOT: 'dashDot',
  DASH_DOT_DOT: 'dashDotDot',
  THICK_THIN: 'double',
  THIN_THICK: 'double',
  TRIPLE: 'double',
};

/** DOCX stroke kind 매핑 */
const DOCX_STROKE: Record<string, StrokeKind> = {
  single: 'solid',
  none: 'none',
  nil: 'none',
  dashed: 'dash',
  dotted: 'dot',
  double: 'double',
  dotDash: 'dashDot',
  dotDotDash: 'dashDotDot',
  thickThin: 'double',
  thinThick: 'double',
  triple: 'double',
  wave: 'wave',
  dashDotStroked: 'dashDot',
  threeDEmboss: 'solid',
  threeDEngrave: 'solid',
};

/** HWPX stroke 값을 Stroke 객체로 변환 */
export function strokeFromHwpx(type?: string, width?: number, color?: string): Stroke {
  return {
    kind: HWPX_STROKE[type ?? ''] ?? 'solid',
    pt: width != null ? width / 100 : 0.5, // hwpunit → pt
    color: safeHex(color) ?? '000000',
  };
}

/** DOCX stroke 값을 Stroke 객체로 변환 */
export function strokeFromDocx(type?: string, size?: number, color?: string): Stroke {
  return {
    kind: DOCX_STROKE[type ?? ''] ?? 'solid',
    pt: size != null ? size / 8 : 0.5, // 1/8 pt 단위로 변환
    color: safeHex(color) ?? '000000',
  };
}

// ─── 스타일 객체 매핑 ───────────────────────────────────

/**
 * DOCX 스타일 → HWPX 스타일 매핑
 */
export function styleDocxToHwpx(docxStyle: Partial<TextProps>): Partial<TextProps> {
  return {
    font: fontToKorean(docxStyle.font),
    pt: docxStyle.pt,
    b: docxStyle.b,
    i: docxStyle.i,
    u: docxStyle.u,
    s: docxStyle.s,
    color: docxStyle.color,
  };
}

/**
 * HWPX 스타일 → DOCX 스타일 매핑
 */
export function styleHwpxToDocx(hwpxStyle: Partial<TextProps>): Partial<TextProps> {
  return {
    font: fontToEnglish(hwpxStyle.font),
    pt: hwpxStyle.pt,
    b: hwpxStyle.b,
    i: hwpxStyle.i,
    u: hwpxStyle.u,
    s: hwpxStyle.s,
    color: hwpxStyle.color,
  };
}

/**
 * DOCX 스타일 → HWP 스타일 매핑
 */
export function styleDocxToHwp(docxStyle: Partial<TextProps>): Partial<TextProps> {
  return {
    font: fontToKorean(docxStyle.font),
    pt: docxStyle.pt,
    b: docxStyle.b,
    i: docxStyle.i,
    u: docxStyle.u,
    s: docxStyle.s,
    color: docxStyle.color,
  };
}

/**
 * HWP 스타일 → DOCX 스타일 매핑
 */
export function styleHwpToDocx(hwpStyle: Partial<TextProps>): Partial<TextProps> {
  return {
    font: fontToEnglish(hwpStyle.font),
    pt: hwpStyle.pt,
    b: hwpStyle.b,
    i: hwpStyle.i,
    u: hwpStyle.u,
    s: hwpStyle.s,
    color: hwpStyle.color,
  };
}
