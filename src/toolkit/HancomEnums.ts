/**
 * 한컴 OWPML 열거형 표.
 *
 * 한컴은 열거 이름을 정확히 일치시켜 읽으므로 표에 있는 이름만 써야 한다.
 * HWP 5 바이너리 값은 배열 인덱스와 같다(선 종류·굵기·textWrap).
 */
import type { ImgWrap, StrokeKind } from '../model/doc-props';

/**
 * BORDER_FILL/hh:*Border type. 인덱스 = HWP 5 바이너리 값 = HWP 5 명세 표 25 순서(+NONE).
 * 이름과 모양이 반대인 항목에 주의: 2 "DOT"은 긴 점선(파선 모양), 3 "DASH"는 점선(점 모양).
 */
export const HANCOM_LINE_TYPES = [
  'NONE', 'SOLID', 'DOT', 'DASH', 'DASH_DOT', 'DASH_DOT_DOT', 'LONG_DASH', 'CIRCLE',
  'DOUBLE_SLIM', 'SLIM_THICK', 'THICK_SLIM', 'SLIM_THICK_SLIM', 'WAVE', 'DOUBLEWAVE',
  'THICK3D', 'THICKREV3D', '3D', 'REV3D',
] as const;

/** BORDER_FILL/hh:*Border width. 인덱스 = HWP 5 바이너리 값. */
export const HANCOM_LINE_WIDTHS = [
  '0.1 mm', '0.12 mm', '0.15 mm', '0.2 mm', '0.25 mm', '0.3 mm', '0.4 mm', '0.5 mm',
  '0.6 mm', '0.7 mm', '1.0 mm', '1.5 mm', '2.0 mm', '3.0 mm', '4.0 mm', '5.0 mm',
] as const;

/** 개체 textWrap. 한컴에는 TIGHT/THROUGH가 없다. 인덱스 = 공통 속성 bit 21-23. */
export const HANCOM_TEXT_WRAPS = ['SQUARE', 'TOP_AND_BOTTOM', 'BEHIND_TEXT', 'IN_FRONT_OF_TEXT'] as const;

const PT_PER_MM = 72 / 25.4;

// IR(DOCX 의미)과 한컴 값은 모양으로 대응한다: dashed = 긴 점선(2), dotted = 점선(3).
const STROKE_TO_LINE_TYPE: Record<StrokeKind, number> = {
  none: 0, solid: 1, dash: 2, dot: 3, dashDot: 4, dashDotDot: 5, double: 8, wave: 12,
};

/** IR 선 종류 → HWP 5 바이너리 선 종류 값. */
export function strokeKindToLineType(kind: StrokeKind | undefined): number {
  return STROKE_TO_LINE_TYPE[kind ?? 'solid'] ?? 1;
}

/** HWP 5 바이너리 선 종류 값 → IR 선 종류. 3D 계열은 실선으로 근사한다. */
export function lineTypeToStrokeKind(code: number): StrokeKind {
  if (code === 0) return 'none';
  if (code === 2 || code === 6) return 'dash';
  if (code === 3 || code === 7) return 'dot';
  if (code === 4) return 'dashDot';
  if (code === 5) return 'dashDotDot';
  if (code >= 8 && code <= 11) return 'double';
  if (code === 12 || code === 13) return 'wave';
  return 'solid';
}

/** IR 선 종류 → OWPML 선 종류 이름. */
export function strokeKindToHwpxLineType(kind: StrokeKind | undefined): string {
  return HANCOM_LINE_TYPES[strokeKindToLineType(kind)];
}

/** OWPML 선 종류 이름 → IR 선 종류. 표에 없는 이름은 undefined. */
export function hwpxLineTypeToStrokeKind(name: string | undefined): StrokeKind | undefined {
  const code = HANCOM_LINE_TYPES.indexOf((name ?? '') as typeof HANCOM_LINE_TYPES[number]);
  return code < 0 ? undefined : lineTypeToStrokeKind(code);
}

/** pt 굵기 → 가장 가까운 한컴 굵기 인덱스(HWP 5 바이너리 값). */
export function nearestLineWidthIndex(pt: number): number {
  const mm = pt / PT_PER_MM;
  let best = 0;
  for (let i = 1; i < HANCOM_LINE_WIDTHS.length; i++) {
    if (Math.abs(parseFloat(HANCOM_LINE_WIDTHS[i]) - mm) < Math.abs(parseFloat(HANCOM_LINE_WIDTHS[best]) - mm)) best = i;
  }
  return best;
}

const WRAP_TO_CODE: Record<ImgWrap, number> = {
  inline: 1, square: 0, tight: 0, through: 0, topAndBottom: 1, behind: 2, front: 3, none: 3,
};

/** IR 배치 → 한컴 textWrap 값(HWP 5 bit 21-23). 한컴에 없는 tight/through는 어울림. */
export function imgWrapToTextWrapCode(wrap: ImgWrap | undefined): number {
  return WRAP_TO_CODE[wrap ?? 'square'] ?? 0;
}

/** IR 배치 → OWPML textWrap 이름. */
export function imgWrapToHwpxTextWrap(wrap: ImgWrap | undefined): string {
  return HANCOM_TEXT_WRAPS[imgWrapToTextWrapCode(wrap)];
}

/** 한컴 textWrap 값 → IR 배치. */
export function textWrapCodeToImgWrap(code: number): ImgWrap {
  return (['square', 'topAndBottom', 'behind', 'front'] as const)[code] ?? 'square';
}

/** IR 쪽 번호 형식 → OWPML NumberType 이름(autoNumFormat type, pageNum formatType). */
export function pageNumFormatToHwpx(format: string | undefined): string {
  return format === 'roman' ? 'ROMAN_SMALL' : format === 'romanCaps' ? 'ROMAN_CAPITAL' : 'DIGIT';
}

/** OWPML NumberType 이름 → IR 쪽 번호 형식. 이전 hwpkit 출력 이름도 읽는다. */
export function hwpxNumberTypeToPageNumFormat(name: string | undefined): 'decimal' | 'roman' | 'romanCaps' {
  if (name === 'ROMAN_SMALL' || name === 'ROMAN_LOWER') return 'roman';
  if (name === 'ROMAN_CAPITAL' || name === 'ROMAN_UPPER') return 'romanCaps';
  return 'decimal';
}
