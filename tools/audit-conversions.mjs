// Usage: node tools/audit-conversions.mjs INPUT_DIRECTORY REPORT.json [md,html,...]
// Reads originals, converts in memory, and writes only the audit report.
import { readdir, readFile, writeFile } from 'node:fs/promises';
import { join, extname } from 'node:path';
import JSZip from 'jszip';
import { SaxesParser } from 'saxes';
import { Pipeline, TreeWalker, registry, BinaryKit } from '../dist/index.mjs';

const [inputDirectory, reportPath, targetList] = process.argv.slice(2);
if (!inputDirectory || !reportPath) throw new Error('Expected INPUT_DIRECTORY REPORT.json');
const outputFormats = ['hwp', 'hwpx', 'docx', 'md', 'html'];
const targets = targetList ? [...new Set(targetList.split(','))] : outputFormats;
if (targets.some(format => !outputFormats.includes(format))) throw new Error('Unsupported audit target');
const inputs = new Set([...outputFormats, 'doc']);
const rows = [];
const walker = new TreeWalker();
const normalized = doc => walker.extractText(doc).replace(/\s/g, '');
const names = (await readdir(inputDirectory)).filter(name => inputs.has(extname(name).slice(1).toLowerCase())).sort();
for (const name of names) {
  const extension = extname(name).slice(1).toLowerCase();
  const data = new Uint8Array(await readFile(join(inputDirectory, name)));
  let format = extension;
  if (data[0] === 0x50 && data[1] === 0x4b) {
    const zip = await JSZip.loadAsync(data);
    format = zip.file('word/document.xml') ? 'docx' : 'hwpx';
  } else if (BinaryKit.isOle2(data)) {
    format = BinaryKit.parseCfb(data).has('WordDocument') ? 'doc' : 'hwp';
  }
  const source = await Pipeline.open(data, ['md', 'html'].includes(format) ? format : undefined).inspect();
  for (const target of targets) {
    const row = { file: name, extension, from: format, to: target, ok: false };
    try {
      if (!source.ok) throw new Error(source.error);
      // Decode each original once; isolate encoder mutations, if any, per output.
      // The separate public API matrix exercises Pipeline.to for all directions.
      const result = await registry.getEncoder(target).encode(structuredClone(source.data));
      if (!result.ok) throw new Error(result.error);
      row.warningCount = result.warns.length;
      if (target === 'hwpx' || target === 'docx') {
        const zip = await JSZip.loadAsync(result.data, { checkCRC32: true });
        for (const entry of Object.values(zip.files)) {
          if (/\.(xml|hpf|rels)$/.test(entry.name)) new SaxesParser({ xmlns: true }).write(await entry.async('string')).close();
        }
      }
      const reopened = await Pipeline.open(result.data, target).inspect();
      if (!reopened.ok) throw new Error(reopened.error);
      const before = normalized(source.data), after = normalized(reopened.data);
      row.sourceCharacters = before.length;
      row.outputCharacters = after.length;
      row.normalizedTextEqual = before === after;
      row.ok = before.length === 0 || after.length > 0;
      if (!row.ok) row.error = 'Nonempty input produced an empty document';
    } catch (error) { row.error = error.message; }
    rows.push(row);
  }
  if ((rows.length / targets.length) % 20 === 0) {
    console.log(`${rows.length / targets.length}/${names.length} files checked; ${rows.filter(r => !r.ok).length} failures`);
    await writeFile(reportPath, JSON.stringify({ inProgress: true, rows }, null, 2) + '\n');
  }
}
const summary = {
  inputFiles: names.length, conversions: rows.length,
  succeeded: rows.filter(r => r.ok).length, failed: rows.filter(r => !r.ok).length,
  normalizedTextEqual: rows.filter(r => r.normalizedTextEqual).length,
  directions: Object.fromEntries([...new Set(rows.map(r => `${r.from}->${r.to}`))].map(direction => {
    const subset = rows.filter(r => `${r.from}->${r.to}` === direction);
    return [direction, { checked: subset.length, failed: subset.filter(r => !r.ok).length,
      normalizedTextEqual: subset.filter(r => r.normalizedTextEqual).length }];
  })),
};
await writeFile(reportPath, JSON.stringify({ summary, rows }, null, 2) + '\n');
console.log(JSON.stringify(summary, null, 2));
if (summary.failed) process.exitCode = 1;
