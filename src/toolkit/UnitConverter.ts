/**
 * 단위 변환 유틸리티
 *
 * 다양한 문서 포맷 간의 단위 변환을 제공합니다.
 *
 * 단위 시스템:
 * - HWP: 1 inch = 7200 hwpunit (1 pt = 100 hwpunit)
 * - DOCX: 1 inch = 1440 dxa = 914400 emu (1 pt = 20 dxa = 12700 emu)
 * - HWPX: charPr height 는 1000 = 10pt
 */

// ─── 기본 단위 변환 (pt 기준) ─────────────────────────────

/** pt 를 다양한 단위로 변환 */
export const ptTo = {
  /** pt → HWP unit (1 pt = 100 hwpunit) */
  hwpunit: (pt: number): number => Math.round(pt * 100),

  /** pt → EMU (English Metric Units, 1 pt = 12700 emu) */
  emu: (pt: number): number => Math.round(pt * 12700),

  /** pt → DXA (Twips, 1 pt = 20 dxa) */
  dxa: (pt: number): number => Math.round(pt * 20),

  /** pt → HWPX charPr height (1 pt = 100 hHeight) */
  hHeight: (pt: number): number => Math.round(pt * 100),

  /** pt → DOCX half-point (1 pt = 2 half-pt) */
  halfPt: (pt: number): number => Math.round(pt * 2),

  /** pt → pixel (96 DPI 기준) */
  pixel: (pt: number): number => Math.round(pt * 1.333),
};

export const mmTo = {
  /** mm → HWP unit (오버플로 방지 및 한컴 엔진 내림 처리 반영) */
  hwpunit: (mm: number): number => Math.floor((mm * 283465) / 1000),
};

/** 다양한 단위를 pt 로 변환 */
export const toPt = {
  /** HWP unit → pt */
  hwpunit: (val: number): number => val / 100,

  /** EMU → pt */
  emu: (val: number): number => val / 12700,

  /** DXA → pt */
  dxa: (val: number): number => val / 20,

  /** HWPX charPr height → pt */
  hHeight: (val: number): number => val / 100,

  /** DOCX half-point → pt */
  halfPt: (val: number): number => val / 2,

  /** pixel → pt (96 DPI 기준) */
  pixel: (val: number): number => val * 0.75,
};

// ─── 직접 변환 (pt 를 거치지 않음) ───────────────────────

/** HWP ↔ DOCX 단위 직접 변환 */
export const directConvert = {
  /** hwpunit → dxa (1 hwpunit = 0.2 dxa) */
  hwpunitToDxa: (val: number): number => Math.round(val / 5),

  /** dxa → hwpunit (1 dxa = 5 hwpunit) */
  dxaToHwpunit: (val: number): number => Math.round(val * 5),

  /** hwpunit → emu (1 hwpunit = 127 emu) */
  hwpunitToEmu: (val: number): number => Math.round(val * 127),

  /** emu → hwpunit (1 emu = 1/127 hwpunit) */
  emuToHwpunit: (val: number): number => Math.round(val / 127),

  /** dxa → emu (1 dxa = 635 emu) */
  dxaToEmu: (val: number): number => Math.round(val * 635),

  /** emu → dxa (1 emu = 1/635 dxa) */
  emuToDxa: (val: number): number => Math.round(val / 635),
};

// ─── 페이지 차원 변환 ────────────────────────────────────

/** A4 크기 (pt) */
export const A4_DIMENSIONS = {
  widthPt: 595.28,
  heightPt: 841.89,
};

/** A4 크기 (hwpunit) */
export const A4_HWPUNIT = {
  width: 5952,
  height: 8418,
};

/** A4 크기 (emu) */
export const A4_EMU = {
  width: 7512960,
  height: 10689180,
};

/** A4 크기 (dxa) */
export const A4_DXA = {
  width: 11905,
  height: 16837,
};

// ─── 편의 함수 ───────────────────────────────────────────

/** pt 를 지정된 포맷의 단위로 변환 */
export function ptToFormat(pt: number, format: 'hwp' | 'docx' | 'hwpx'): number {
  switch (format) {
    case 'hwp':
      return ptTo.hwpunit(pt);
    case 'docx':
      return ptTo.emu(pt);
    case 'hwpx':
      return ptTo.hHeight(pt);
    default:
      return pt;
  }
}

/** 지정된 포맷의 단위를 pt 로 변환 */
export function formatToPt(value: number, format: 'hwp' | 'docx' | 'hwpx'): number {
  switch (format) {
    case 'hwp':
      return toPt.hwpunit(value);
    case 'docx':
      return toPt.emu(value);
    case 'hwpx':
      return toPt.hHeight(value);
    default:
      return value;
  }
}
