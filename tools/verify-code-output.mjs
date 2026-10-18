/** Verify the built public ESM/CommonJS APIs without a browser or Canvas.
 * node tools/verify-code-output.mjs SOURCE.docx OUTPUT_DIRECTORY
 */
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { createRequire } from 'node:module';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';

const [sourcePath, outputArg] = process.argv.slice(2);
if (!sourcePath || !outputArg) throw Error('Usage: node tools/verify-code-output.mjs SOURCE.docx OUTPUT_DIRECTORY');
const output = resolve(outputArg);
await mkdir(output, { recursive: true });
// Resolve both public package export conditions from inside this package.
const esm = await import('hwpkit-dev');
const cjs = createRequire(import.meta.url)('hwpkit-dev');
const input = await readFile(sourcePath);
const report = { source: resolve(sourcePath), sourceSha256: hash(input), node: process.version,
  canvasAvailable: typeof OffscreenCanvas !== 'undefined', cases: [] };
const unpack = r => { if (!r.ok) throw Error(r.error); return r.data; };
function hash(bytes) { return createHash('sha256').update(bytes).digest('hex'); }
function walk(node, callback) { callback(node); for (const kid of node.kids ?? []) walk(kid, callback); }
function glyphs(doc) {
  const out = [];
  walk(doc, n => { if (n.tag === 'span') for (const t of n.kids) if (t.tag === 'txt')
    for (const ch of t.content) if (!/\s/u.test(ch)) out.push({ ch, b: !!n.props.b, i: !!n.props.i }); });
  return out;
}
function images(doc) { const out = []; walk(doc, n => { if (n.tag === 'img') out.push(n); }); return out; }
function close(actual, expected, label) { assert.ok(Math.abs(actual - expected) <= .03, `${label}: ${actual} != ${expected}`); }
function checkDocument(original, decoded) {
  const a = glyphs(original), b = glyphs(decoded);
  assert.deepEqual(b, a, 'text or per-character emphasis changed');
  assert.equal(decoded.kids.length, original.kids.length, 'section count');
  for (let i = 0; i < original.kids.length; i++) for (const key of ['wPt', 'hPt', 'mt', 'mb', 'ml', 'mr'])
    close(decoded.kids[i].dims[key], original.kids[i].dims[key], `section ${i} ${key}`);
  const before = images(original), after = images(decoded);
  assert.equal(after.length, before.length, 'image count');
  for (let i = 0; i < before.length; i++) {
    const a = before[i], b = after[i];
    close(b.w, a.w, `image ${i} width`); close(b.h, a.h, `image ${i} height`);
    assert.ok(b.b64.length > 0, `image ${i} empty`);
    assert.equal(b.b64, a.b64, `image ${i} bytes changed`);
    for (const key of ['wrap', 'horzRelTo', 'vertRelTo', 'horzAlign', 'vertAlign'])
      if (a.layout?.[key] !== undefined) assert.equal(b.layout?.[key], a.layout[key], `image ${i} ${key}`);
    for (const key of ['xPt', 'yPt']) if (a.layout?.[key] !== undefined)
      close(b.layout?.[key], a.layout[key], `image ${i} ${key}`);
  }
  return { characters: a.length, emphasisMismatches: 0, images: after.map(({ w, h, layout, b64 }) =>
    ({ w, h, layout, byteCount: Buffer.from(b64, 'base64').length })), page: decoded.kids[0].dims };
}

for (const [moduleFormat, kit] of [['esm', esm], ['cjs', cjs]]) {
  const source = unpack(await kit.Pipeline.open(input, 'docx').inspect());
  const autoSource = unpack(await (await kit.Pipeline.openAsync(new Blob([input]))).inspect());
  assert.deepEqual(glyphs(autoSource), glyphs(source), 'Blob automatic input detection');
  for (const format of ['hwp', 'hwpx']) {
    const converted = await kit.Pipeline.open(input, 'docx').to(format);
    const bytes = unpack(converted);
    const autoBytes = unpack(await kit.Pipeline.open(input).to('.' + format.toUpperCase()));
    // ZIP timestamps may differ; compare decoded content instead of entire ZIPs.
    const decoded = unpack(await kit.Pipeline.open(bytes).inspect());
    const checks = checkDocument(source, decoded);
    checkDocument(source, unpack(await kit.Pipeline.open(autoBytes).inspect()));
    const blobBytes = unpack(await (await kit.Pipeline.openAsync(new Blob([input]))).to(format));
    checkDocument(source, unpack(await kit.Pipeline.open(blobBytes).inspect()));
    const path = join(output, `file-sample_1MB-${moduleFormat}.${format}`);
    await writeFile(path, bytes);
    report.cases.push({ moduleFormat, inputMode: 'Buffer / automatic / async Blob', source: 'docx', format,
      path, bytes: bytes.length, sha256: hash(bytes), warns: converted.warns, ...checks });
  }

  // Also exercise documents assembled directly in code, including Korean,
  // explicit emphasis OFF, script-specific faces, an inline image and a table.
  const { buildRoot, buildSheet, buildPara, buildSpan, buildImg, buildGrid, buildRow, buildCell } = kit;
  const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
  const face = { font: '바탕', fontFaces: { hangul: '바탕', latin: 'Arial' }, pt: 11 };
  const model = buildRoot({}, [buildSheet([
    buildPara([buildSpan('일반 Regular ', { ...face, b: false, i: false }), buildSpan('진하게 Bold ', { ...face, b: true }),
      buildSpan('기울임 Italic ', { ...face, i: true }), buildSpan('다시 일반 Normal', { ...face, b: false, i: false })],
      { indentPt: 36, firstLineIndentPt: -18, spaceAfter: 0 }),
    buildPara([buildSpan('그림 앞 ', face), buildImg(png, 'image/png', 24, 16), buildSpan(' 그림 뒤', face)]),
    buildGrid([buildRow([buildCell([buildPara([buildSpan('표 안의 일반 문장 ABC', face)])])])]),
  ], { wPt: 595.3, hPt: 841.9, mt: 56.7, mb: 56.7, ml: 56.7, mr: 56.7 })]);
  const fixtureDocx = unpack(await kit.registry.getEncoder('docx').encode(model));
  await writeFile(join(output, `code-model-${moduleFormat}.docx`), fixtureDocx);
  for (const format of ['hwp', 'hwpx']) {
    const converted = await kit.registry.getEncoder(format).encode(model);
    const bytes = unpack(converted), decoded = unpack(await kit.Pipeline.open(bytes).inspect());
    const checks = checkDocument(model, decoded);
    let spanCount = 0;
    walk(decoded, n => { if (n.tag === 'span' && n.kids.some(t => t.tag === 'txt' && t.content.trim())) {
      assert.equal(n.props.fontFaces?.hangul, '바탕'); assert.equal(n.props.fontFaces?.latin, 'Arial');
      close(n.props.pt, 11, 'run font size'); spanCount++;
    } });
    assert.ok(spanCount >= 6);
    const path = join(output, `code-model-${moduleFormat}.${format}`); await writeFile(path, bytes);
    report.cases.push({ moduleFormat, source: 'code model', format, path, bytes: bytes.length,
      sha256: hash(bytes), warns: converted.warns, ...checks });
  }
}
await writeFile(join(output, 'code-api-report.json'), JSON.stringify(report, null, 2) + '\n');
console.log(JSON.stringify({ casesPassed: report.cases.length, canvasAvailable: report.canvasAvailable,
  report: join(output, 'code-api-report.json'), outputs: report.cases.map(c => ({ format: c.format, module: c.moduleFormat,
    source: c.source, characters: c.characters, bytes: c.bytes, warns: c.warns })) }, null, 2));
