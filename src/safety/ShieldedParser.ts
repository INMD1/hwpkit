export class ShieldedParser {
  private log: string[] = [];

  /** 단일 요소 안전 파싱 */
  guard<T>(fn: () => T, fallback: T, label: string): T {
    try {
      const v = fn();
      if (v == null) {
        this.warn(label, 'returned null/undefined');
        return fallback;
      }
      return v;
    } catch (e: any) {
      this.warn(label, e?.message ?? String(e));
      return fallback;
    }
  }

  /** 배열 각 요소 독립 파싱 (하나 실패해도 나머지 계속) */
  guardAll<I, O>(
    items: I[],
    fn: (x: I, i: number) => O,
    fb: (x: I, i: number) => O,
    label: string,
  ): O[] {
    return items.map((x, i) =>
      this.guard(() => fn(x, i), fb(x, i), `${label}[${i}]`),
    );
  }

  /**
   * 표 전용 4단계 폴백
   *   Lv1: Full → Lv2: Grid → Lv3: Flat → Lv4: Text
   */
  guardGrid<T>(
    node: unknown,
    lv1Full: (n: unknown) => T,
    lv2Grid: (n: unknown) => T,
    lv3Flat: (n: unknown) => T,
    lv4Text: (n: unknown) => T,
    label: string,
  ): { value: T; level: 1 | 2 | 3 | 4 } {
    const levels: [(n: unknown) => T, 1 | 2 | 3 | 4][] = [
      [lv1Full, 1], [lv2Grid, 2], [lv3Flat, 3], [lv4Text, 4],
    ];

    for (const [fn, lv] of levels) {
      try {
        const v = fn(node);
        if (v != null) {
          if (lv > 1) this.warn(label, `degraded to level ${lv}`);
          return { value: v, level: lv };
        }
      } catch (e: any) {
        this.warn(label, `Lv${lv} failed: ${e?.message ?? String(e)}`);
      }
    }

    this.warn(label, 'ALL LEVELS FAILED — returning lv4Text forced');
    return { value: lv4Text(null), level: 4 };
  }

  /** 이미지 안전 파싱 */
  guardImg<T>(
    node: unknown,
    fn: (n: unknown) => T,
    placeholder: (alt: string) => T,
    label: string,
  ): T {
    try {
      const v = fn(node);
      if (v != null) return v;
    } catch (e: any) {
      this.warn(label, e?.message ?? String(e));
    }
    this.warn(label, 'using placeholder image');
    return placeholder(`[이미지 로드 실패: ${label}]`);
  }

  private warn(label: string, msg: string): void {
    const w = `[SHIELD] ${label}: ${msg}`;
    console.warn(w);
    this.log.push(w);
  }

  flush(): string[] {
    const r = [...this.log];
    this.log = [];
    return r;
  }
}
