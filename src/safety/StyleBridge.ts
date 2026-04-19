import type { Align, StrokeKind, Stroke } from '../model/doc-props';

// ─── 단위 변환 ─────────────────────────────────────────────
export const Metric = {
  // HWP 세계 (1 inch = 7200 HWPUNIT)
  hwpToPt:     (v: number) => v / 100,
  ptToHwp:     (v: number) => Math.round(v * 100),
  hwpToDxa:    (v: number) => Math.round(v / 5),
  dxaToHwp:    (v: number) => Math.round(v * 5),
  hwpToEmu:    (v: number) => Math.round(v * 127),
  emuToHwp:    (v: number) => Math.round(v / 127),

  // DOCX 세계 (1 inch = 1440 dxa, 1 pt = 20 dxa)
  dxaToPt:     (v: number) => v / 20,
  ptToDxa:     (v: number) => Math.round(v * 20),
  dxaToEmu:    (v: number) => Math.round(v * 635),
  emuToDxa:    (v: number) => Math.round(v / 635),
  emuToPt:     (v: number) => v / 12700,
  ptToEmu:     (v: number) => Math.round(v * 12700),

  // HWPX charPr height: 1000 = 10pt
  hHeightToPt: (v: number) => v / 100,
  ptToHHeight: (v: number) => Math.round(v * 100),

  // DOCX half-point: 24 = 12pt
  halfPtToPt:  (v: number) => v / 2,
  ptToHalfPt:  (v: number) => Math.round(v * 2),
} as const;

// ─── 색상 정규화 ───────────────────────────────────────────
export function safeHex(raw: string | number | null | undefined): string | undefined {
  if (raw == null) return undefined;
  if (typeof raw === 'number') {
    if (raw <= 0) return '000000';
    if (raw >= 0xFFFFFF) return undefined;
    return raw.toString(16).padStart(6, '0').toUpperCase();
  }
  let s = String(raw).replace(/^#/, '').toUpperCase();
  if (/^[0-9A-F]{3}$/.test(s)) s = s[0] + s[0] + s[1] + s[1] + s[2] + s[2];
  if (/^[0-9A-F]{6}$/.test(s)) return s;
  if (s === 'AUTO' || s === 'NONE' || s === 'TRANSPARENT') return undefined;
  return undefined;
}

// ─── 정렬 정규화 ───────────────────────────────────────────
const ALIGN_MAP: Record<string, Align> = {
  LEFT: 'left', CENTER: 'center', RIGHT: 'right', JUSTIFY: 'justify',
  BOTH: 'justify', DISTRIBUTE: 'justify',
  left: 'left', center: 'center', right: 'right', both: 'justify',
  start: 'left', end: 'right',
};
export function safeAlign(raw?: string): Align {
  return ALIGN_MAP[raw ?? ''] ?? 'left';
}

// ─── 테두리 정규화 ─────────────────────────────────────────
const HWPX_STROKE: Record<string, StrokeKind> = {
  SOLID: "solid",
  NONE: "none",
  DASH: "dash",
  DOT: "dot",
  DOUBLE: "double",
  LONG_DASH: "dash",
  DASH_DOT: "dashDot",
  DASH_DOT_DOT: "dashDotDot",
  THICK_THIN: "double",
  THIN_THICK: "double",
  TRIPLE: "double",
};
const DOCX_STROKE: Record<string, StrokeKind> = {
  single: "solid",
  none: "none",
  nil: "none",
  dashed: "dash",
  dotted: "dot",
  double: "double",
  dotDash: "dashDot",
  dotDotDash: "dashDotDot",
  thickThin: "double",
  thinThick: "double",
  triple: "double",
  wave: "wave",
  dashDotStroked: "dashDot",
  threeDEmboss: "solid",
  threeDEngrave: "solid",
};

export function safeStrokeHwpx(type?: string, w?: number, c?: string): Stroke {
  return {
    kind: HWPX_STROKE[type ?? ''] ?? 'solid',
    pt: w != null ? Metric.hwpToPt(w) : 0.5,
    color: safeHex(c) ?? '000000',
  };
}

export function safeStrokeDocx(val?: string, sz?: number, c?: string): Stroke {
  return {
    kind: DOCX_STROKE[val ?? ''] ?? 'solid',
    pt: sz != null ? sz / 8 : 0.5,
    color: safeHex(c) ?? '000000',
  };
}

// ─── 폰트 정규화 ───────────────────────────────────────────
const FONT_MAP: Record<string, string> = {
  '맑은 고딕': 'Malgun Gothic',
  '바탕': 'Batang',
  '돋움': 'Dotum',
  '굴림': 'Gulim',
  '한컴바탕': 'Batang',
  '한컴돋움': 'Malgun Gothic',
  '함초롬바탕': 'Batang',
  '함초롬돋움': 'Malgun Gothic',
};
export function safeFont(raw?: string): string {
  return FONT_MAP[raw ?? ''] ?? raw ?? 'Malgun Gothic';
}

// Reverse mapping: English → Korean (for HWPX encoding)
const FONT_MAP_KR: Record<string, string> = {
  'Malgun Gothic': '맑은 고딕',
  'Batang': '바탕',
  'Dotum': '돋움',
  'Gulim': '굴림',
};
export function safeFontToKr(raw?: string): string {
  return FONT_MAP_KR[raw ?? ''] ?? raw ?? '맑은 고딕';
}
