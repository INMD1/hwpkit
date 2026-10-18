/**
 * A width observation taken from a source table cell.
 *
 * The source formats store the width of a merged cell, not always the width of
 * every underlying column.  Each observation therefore represents the linear
 * constraint:
 *
 *   column[start] + ... + column[start + span - 1] = width
 */
export interface ColumnWidthConstraint {
  start: number;
  span: number;
  width: number;
}

/**
 * Normalize integer column widths so a table never exceeds its container.
 *
 * Missing columns consume the remaining space, while over-wide tables are
 * reduced proportionally. Largest-remainder rounding keeps the result stable
 * and guarantees that the integer sum is no greater than `maxTotal`.
 */
export function fitColumnWidths(
  source: readonly number[],
  columnCount: number,
  maxTotal: number,
  minWidth = 1,
): number[] {
  const n = Math.max(0, Math.floor(columnCount));
  if (n === 0) return [];

  const limit = Math.max(n, Math.floor(Number.isFinite(maxTotal) ? maxTotal : n));
  const floorWidth = Math.max(1, Math.min(Math.floor(minWidth), Math.floor(limit / n)));
  const values = Array.from({ length: n }, (_, index) => {
    const value = Number(source[index]);
    return Number.isFinite(value) && value > 0 ? value : 0;
  });

  const knownTotal = values.reduce((sum, value) => sum + value, 0);
  const missing = values.reduce((count, value) => count + (value <= 0 ? 1 : 0), 0);
  if (missing === n) {
    values.fill(limit / n);
  } else if (missing > 0) {
    const remaining = limit - knownTotal;
    const knownAverage = knownTotal / (n - missing);
    const fill = remaining > 0 ? remaining / missing : knownAverage;
    for (let i = 0; i < n; i++) if (values[i] <= 0) values[i] = Math.max(1, fill);
  }

  const rawTotal = values.reduce((sum, value) => sum + value, 0);
  const target = Math.max(
    floorWidth * n,
    Math.min(limit, Math.round(rawTotal > 0 ? rawTotal : limit)),
  );
  const exact = new Array(n).fill(0);
  const active = new Set(Array.from({ length: n }, (_, index) => index));
  let remaining = target;
  while (active.size > 0) {
    const weightTotal = [...active].reduce((sum, index) => sum + values[index], 0);
    const tooSmall = [...active].filter((index) =>
      weightTotal <= 0 || (remaining * values[index]) / weightTotal < floorWidth,
    );
    if (tooSmall.length === 0) {
      for (const index of active) exact[index] = (remaining * values[index]) / weightTotal;
      break;
    }
    for (const index of tooSmall) {
      exact[index] = floorWidth;
      remaining -= floorWidth;
      active.delete(index);
    }
  }
  const result = exact.map(Math.floor);
  let residual = target - result.reduce((sum, value) => sum + value, 0);
  const order = exact
    .map((value, index) => ({ index, fraction: value - Math.floor(value) }))
    .sort((a, b) => b.fraction - a.fraction || a.index - b.index);
  for (let i = 0; residual > 0; i++, residual--) {
    result[order[i % order.length].index]++;
  }
  return result;
}

/**
 * Infer physical column widths from all merged and unmerged cells at once.
 *
 * A ridge-regularised least-squares solve is used instead of dividing every
 * merged cell equally.  The small prior only selects a stable solution when a
 * document does not contain enough independent constraints; it has negligible
 * effect for a fully constrained table.  A projection pass then removes the
 * remaining floating-point residual while keeping every column positive.
 */
export function inferColumnWidths(
  columnCount: number,
  observations: ColumnWidthConstraint[],
): number[] {
  const n = Math.max(0, Math.floor(columnCount));
  if (n === 0) return [];

  const constraints = observations
    .map((item) => ({
      start: Math.max(0, Math.floor(item.start)),
      span: Math.max(1, Math.floor(item.span)),
      width: Number(item.width),
    }))
    .filter(
      (item) =>
        Number.isFinite(item.width) &&
        item.width > 0 &&
        item.start < n &&
        item.start + item.span <= n,
    );

  if (constraints.length === 0) return new Array(n).fill(0);

  const priorSums = new Array(n).fill(0);
  const priorWeights = new Array(n).fill(0);
  for (const item of constraints) {
    const estimate = item.width / item.span;
    const weight = 1 / item.span;
    for (let col = item.start; col < item.start + item.span; col++) {
      priorSums[col] += estimate * weight;
      priorWeights[col] += weight;
    }
  }

  const fullWidths = constraints
    .filter((item) => item.start === 0 && item.span === n)
    .map((item) => item.width)
    .sort((a, b) => a - b);
  const totalHint = fullWidths.length
    ? fullWidths[Math.floor(fullWidths.length / 2)]
    : constraints.reduce((largest, item) => Math.max(largest, item.width), 0);
  const defaultWidth = totalHint / n;
  const prior = priorSums.map((sum, col) =>
    priorWeights[col] > 0 ? sum / priorWeights[col] : defaultWidth,
  );

  // Word tables are normally far below this size.  Keep damaged/adversarial
  // files from allocating an O(n^2) matrix while still returning a usable
  // proportional fallback.
  if (n > 256) {
    const priorTotal = prior.reduce((sum, width) => sum + Math.max(0, width), 0);
    const scale = priorTotal > 0 ? totalHint / priorTotal : 1;
    return prior.map((width) => Math.max(1e-4, width * scale));
  }

  // Normal equations: (A^T A + lambda I)x = A^T b + lambda*prior.
  const matrix = Array.from({ length: n }, () => new Array(n).fill(0));
  const rhs = new Array(n).fill(0);
  for (const item of constraints) {
    // Shorter spans contain more column-specific information.
    const weight = 1 / Math.sqrt(item.span);
    for (let a = item.start; a < item.start + item.span; a++) {
      rhs[a] += weight * item.width;
      for (let b = item.start; b < item.start + item.span; b++) {
        matrix[a][b] += weight;
      }
    }
  }

  const maxDiagonal = Math.max(1, ...matrix.map((row, i) => row[i]));
  const ridge = maxDiagonal * 1e-8;
  for (let i = 0; i < n; i++) {
    matrix[i][i] += ridge;
    rhs[i] += ridge * prior[i];
  }

  const solved = solveLinearSystem(matrix, rhs);
  const minWidth = Math.max(1e-4, totalHint * 1e-8);
  const widths = (solved ?? prior).map((value, col) =>
    Number.isFinite(value) && value > minWidth
      ? value
      : Math.max(minWidth, prior[col]),
  );

  // Kaczmarz-style projections make consistent source constraints exact.  If
  // the source contains rounding conflicts this converges to a stable average.
  for (let pass = 0; pass < 80; pass++) {
    let maxRelativeResidual = 0;
    for (const item of constraints) {
      let actual = 0;
      for (let col = item.start; col < item.start + item.span; col++) {
        actual += widths[col];
      }
      const residual = item.width - actual;
      maxRelativeResidual = Math.max(
        maxRelativeResidual,
        Math.abs(residual) / Math.max(1, item.width),
      );
      if (Math.abs(residual) <= 1e-9) continue;

      const adjustable = widths
        .slice(item.start, item.start + item.span)
        .reduce((sum, value) => sum + Math.max(minWidth, value), 0);
      for (let col = item.start; col < item.start + item.span; col++) {
        const share = Math.max(minWidth, widths[col]) / adjustable;
        widths[col] = Math.max(minWidth, widths[col] + residual * share);
      }
    }
    if (maxRelativeResidual < 1e-8) break;
  }

  return widths;
}

function solveLinearSystem(matrix: number[][], rhs: number[]): number[] | null {
  const n = rhs.length;
  const augmented = matrix.map((row, i) => [...row, rhs[i]]);

  for (let col = 0; col < n; col++) {
    let pivot = col;
    for (let row = col + 1; row < n; row++) {
      if (Math.abs(augmented[row][col]) > Math.abs(augmented[pivot][col])) {
        pivot = row;
      }
    }
    if (Math.abs(augmented[pivot][col]) < 1e-12) return null;
    if (pivot !== col) [augmented[pivot], augmented[col]] = [augmented[col], augmented[pivot]];

    const divisor = augmented[col][col];
    for (let j = col; j <= n; j++) augmented[col][j] /= divisor;

    for (let row = 0; row < n; row++) {
      if (row === col) continue;
      const factor = augmented[row][col];
      if (Math.abs(factor) < 1e-15) continue;
      for (let j = col; j <= n; j++) {
        augmented[row][j] -= factor * augmented[col][j];
      }
    }
  }

  return augmented.map((row) => row[n]);
}
