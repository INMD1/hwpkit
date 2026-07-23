import { describe, expect, it } from "vitest";
import { fitColumnWidths, inferColumnWidths } from "./TableGeometry";

describe("fitColumnWidths", () => {
  it("proportionally fits an oversized table to its container", () => {
    const widths = fitColumnWidths([6000, 3000, 1000], 3, 5000, 100);

    expect(widths).toEqual([3000, 1500, 500]);
    expect(widths.reduce((sum, width) => sum + width, 0)).toBe(5000);
  });

  it("fills missing columns without exceeding the exact integer limit", () => {
    const widths = fitColumnWidths([1200, 0, 300], 3, 2401, 100);

    expect(widths).toHaveLength(3);
    expect(widths.every((width) => width >= 100)).toBe(true);
    expect(widths.reduce((sum, width) => sum + width, 0)).toBe(2401);
  });

  it("keeps a complete table narrower than the container unchanged", () => {
    expect(fitColumnWidths([400, 300], 2, 1000, 10)).toEqual([400, 300]);
  });
});

describe("inferColumnWidths", () => {
  it("recovers columns from overlapping merged-cell constraints", () => {
    const widths = inferColumnWidths(10, [
      { start: 0, span: 2, width: 8034 },
      { start: 2, span: 1, width: 8034 },
      { start: 3, span: 3, width: 15941 },
      { start: 6, span: 3, width: 7735 },
      { start: 9, span: 1, width: 7734 },
      { start: 0, span: 10, width: 47478 },
      { start: 0, span: 1, width: 6014 },
      { start: 1, span: 3, width: 15391 },
      { start: 4, span: 1, width: 5104 },
      { start: 5, span: 5, width: 20969 },
      { start: 6, span: 2, width: 5911 },
      { start: 8, span: 2, width: 9558 },
    ]);

    expect(widths).toHaveLength(10);
    expect(widths.reduce((sum, width) => sum + width, 0)).toBeCloseTo(47478, 3);
    expect(widths[0]).toBeCloseTo(6014, 3);
    expect(widths[2]).toBeCloseTo(8034, 3);
    expect(widths[4]).toBeCloseTo(5104, 3);
    expect(widths[9]).toBeCloseTo(7734, 3);
  });

  it("keeps underdetermined columns positive", () => {
    const widths = inferColumnWidths(3, [
      { start: 0, span: 3, width: 300 },
      { start: 0, span: 1, width: 60 },
    ]);

    expect(widths[0]).toBeCloseTo(60, 3);
    expect(widths[1]).toBeGreaterThan(0);
    expect(widths[2]).toBeGreaterThan(0);
    expect(widths.reduce((sum, width) => sum + width, 0)).toBeCloseTo(300, 3);
  });
});
