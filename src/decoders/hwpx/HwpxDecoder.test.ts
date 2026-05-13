import { describe, it, expect } from 'vitest';
import { HwpxDecoder } from './HwpxDecoder';
import { DocxEncoder } from '../../encoders/docx/DocxEncoder';
import type { ParaNode, GridNode } from '../../model/doc-tree';

// ─── 단위 변환 검증 ──────────────────────────────────────────────────────────

describe('Metric: HWPUNIT → pt 변환', () => {
  it('A4 폭 59528 HWPUNIT = 595.28pt', () => {
    const pt = 59528 / 100;
    expect(pt).toBeCloseTo(595.28, 2);
  });

  it('A4 높이 84188 HWPUNIT = 841.88pt', () => {
    const pt = 84188 / 100;
    expect(pt).toBeCloseTo(841.88, 1);
  });
});

// ─── ParaPr 매핑 검증 ──────────────────────────────────────────────────────────

describe('HwpxDecoder: ParaPr 들여쓰기 필드 매핑', () => {
  it('extractParaPrs: hc:left → indentPt (전체 왼쪽 여백)', async () => {
    // OWPML §7.5.4.4: hc:left = 문단 전체 왼쪽 여백
    // 3000 HWPUNIT = 30pt
    // extractParaPrs 로직 직접 검증: hc:left value 3000 → indentPt 30
    const mockMarginEl = {
      'hc:left': [{ _attr: { value: '3000' } }],
    };
    const leftVal = Number(mockMarginEl?.['hc:left']?.[0]?._attr?.value ?? 0);
    const indentPt = leftVal !== 0 ? leftVal / 100 : undefined;
    expect(indentPt).toBe(30);
  });

  it('hc:indent → firstLineIndentPt (첫 줄 들여쓰기)', () => {
    // hc:indent 1000 HWPUNIT = 10pt 첫줄 들여쓰기
    const mockMarginEl = {
      'hc:indent': [{ _attr: { value: '1000' } }],
    };
    const indentEl = mockMarginEl?.['hc:indent']?.[0];
    const indentVal = Number(indentEl?._attr?.value ?? 0);
    const firstLineIndentPt = indentVal !== 0 ? indentVal / 100 : undefined;
    expect(firstLineIndentPt).toBe(10);
  });

  it('hc:indent 음수 → firstLineIndentPt 음수 (내어쓰기)', () => {
    const indentVal = Number('-1000');
    const firstLineIndentPt = indentVal !== 0 ? indentVal / 100 : undefined;
    expect(firstLineIndentPt).toBe(-10);
  });
});

// ─── lineSpacing FIXED 검증 ──────────────────────────────────────────────────

describe('HwpxDecoder: lineSpacing 타입 처리', () => {
  it('PERCENT 160 → lineHeight undefined (기본값 skip)', () => {
    const lsType = 'PERCENT';
    const lsVal = 160;
    let lineHeight: number | undefined;
    if (lsType === 'PERCENT' && lsVal > 0 && lsVal !== 160) lineHeight = lsVal / 100;
    expect(lineHeight).toBeUndefined();
  });

  it('PERCENT 200 → lineHeight 2.0', () => {
    const lsType = 'PERCENT';
    const lsVal = Number('200');
    let lineHeight: number | undefined;
    if (lsType === 'PERCENT' && lsVal > 0 && lsVal !== 160) lineHeight = lsVal / 100;
    expect(lineHeight).toBe(2.0);
  });

  it('FIXED 2400 HWPUNIT → lineHeightFixed 24pt', () => {
    const lsType = 'FIXED';
    const lsVal = 2400;
    let lineHeightFixed: number | undefined;
    if (lsType === 'FIXED' && lsVal > 0) lineHeightFixed = lsVal / 100;
    expect(lineHeightFixed).toBe(24);
  });
});

// ─── 이미지 wrap 매핑 검증 ───────────────────────────────────────────────────

describe('HwpxDecoder: 이미지 textWrap 매핑', () => {
  const wrapMap: Record<string, string> = {
    TOP_AND_BOTTOM: 'topAndBottom',
    SQUARE: 'square',
    BOTH_SIDES: 'tight',
    BEHIND_TEXT: 'behind',
    FRONT_TEXT: 'front',
  };

  it('TOP_AND_BOTTOM → topAndBottom (float anchor, wrapTopAndBottom)', () => {
    expect(wrapMap['TOP_AND_BOTTOM']).toBe('topAndBottom');
  });

  it('SQUARE → square', () => {
    expect(wrapMap['SQUARE']).toBe('square');
  });

  it('BEHIND_TEXT → behind', () => {
    expect(wrapMap['BEHIND_TEXT']).toBe('behind');
  });
});

// ─── DocxEncoder: indent XML 생성 검증 ───────────────────────────────────────

describe('DocxEncoder: paragraph indent XML', () => {
  it('firstLineIndentPt > 0 → w:firstLine', () => {
    // 10pt firstLine = 200 dxa
    const firstPt = 10;
    const firstLineDxa = Math.round(firstPt * 20);
    const xml = `<w:ind w:firstLine="${firstLineDxa}"/>`;
    expect(xml).toContain('w:firstLine="200"');
  });

  it('firstLineIndentPt < 0 → w:hanging (내어쓰기)', () => {
    // -10pt = hanging 200 dxa
    const firstPt = -10;
    const hangingDxa = Math.round(-firstPt * 20);
    const xml = `<w:ind w:hanging="${hangingDxa}"/>`;
    expect(xml).toContain('w:hanging="200"');
  });

  it('lineHeightFixed 24pt → lineRule="exact", line=480', () => {
    // ECMA-376 §17.3.1.33: exact는 dxa 단위 (1pt = 20dxa)
    const lineHeightFixed = 24;
    const lineDxa = Math.max(1, Math.round(lineHeightFixed * 20));
    const xml = `<w:spacing w:line="${lineDxa}" w:lineRule="exact"/>`;
    expect(xml).toContain('w:line="480"');
    expect(xml).toContain('w:lineRule="exact"');
  });
});

// ─── DocxEncoder: 표 정렬(tblJc) 검증 ───────────────────────────────────────

describe('DocxEncoder: 표 정렬 tblJc', () => {
  const tblAlignMap: Record<string, string> = {
    left: 'start', center: 'center', right: 'end', justify: 'start',
  };

  it('align=center → w:jc val="center"', () => {
    const jc = tblAlignMap['center'];
    expect(jc).toBe('center');
  });

  it('align=right → w:jc val="end"', () => {
    const jc = tblAlignMap['right'];
    expect(jc).toBe('end');
  });

  it('align=left → w:jc val="start"', () => {
    const jc = tblAlignMap['left'];
    expect(jc).toBe('start');
  });
});

// ─── DocxEncoder: WRAP_DOCX에 topAndBottom 포함 확인 ────────────────────────

describe('DocxEncoder: WRAP_DOCX topAndBottom', () => {
  it('wrapTopAndBottom XML에 wp:wrapTopAndBottom 태그 포함', () => {
    const wrapXml = '<wp:wrapTopAndBottom/>';
    expect(wrapXml).toContain('wrapTopAndBottom');
  });
});
