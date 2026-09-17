import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { createLibreOfficeDocConverter } from './index';
import { configureDocConverter, Pipeline, TreeWalker } from '../index';

const work: string[] = [];
afterEach(async () => {
  configureDocConverter();
  for (const path of work.splice(0)) await rm(path, { recursive: true, force: true });
});

async function fakeOffice(body: string) {
  const directory = await mkdtemp(join(tmpdir(), 'hwpkit-office-test-'));
  work.push(directory);
  const executable = join(directory, 'office with spaces.cjs');
  await writeFile(executable, `#!${process.execPath}\n${body}`, { mode: 0o755 });
  return executable;
}

describe('local DOC bridge failure handling', () => {
  it('reports a missing executable', async () => {
    const convert = createLibreOfficeDocConverter({ executable: '/nonexistent/hwpkit-office' });
    await expect(convert(new Uint8Array())).rejects.toThrow('LibreOffice 실행 실패');
  });
  it('does not accept exit code 0 without an output file', async () => {
    const executable = await fakeOffice('process.exit(0);');
    await expect(createLibreOfficeDocConverter({ executable })(new Uint8Array())).rejects.toThrow('DOCX를 생성하지 못했습니다');
  });
  it('terminates a hung conversion', async () => {
    const executable = await fakeOffice('setInterval(() => {}, 1000);');
    await expect(createLibreOfficeDocConverter({ executable, timeoutMs: 100 })(new Uint8Array())).rejects.toThrow('시간 초과');
  });
  it('removes temporary input/profile files after conversion failure', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'hwpkit-office-record-'));
    work.push(directory);
    const report = join(directory, 'path.txt');
    const executable = await fakeOffice(`require('node:fs').writeFileSync(${JSON.stringify(report)}, process.argv.at(-1)); process.exit(1);`);
    await expect(createLibreOfficeDocConverter({ executable })(new Uint8Array())).rejects.toThrow('변환 실패');
    const input = await readFile(report, 'utf8');
    await expect(readFile(input)).rejects.toThrow();
  });
});

// Run explicitly when local LibreOffice is available; no network or user files.
describe.skipIf(process.env.HWPKIT_TEST_LIBREOFFICE !== '1')('real LibreOffice DOC conversion', () => {
  it('preserves a table and Korean text across DOC → HWP/HWPX/DOCX', async () => {
    const input = new Uint8Array(await readFile(resolve('tests/fixtures/legacy-word97.doc')));
    const bridge = createLibreOfficeDocConverter();
    // Cache the office normalization once; each target still decodes it independently.
    const docx = await bridge(input);
    configureDocConverter(async () => docx);
    for (const format of ['hwp', 'hwpx', 'docx']) {
      const result = await Pipeline.open(input).to(format);
      if (!result.ok) throw new Error(result.error);
      expect(result.warns.join(' ')).not.toContain('기본 읽기');
      const inspected = await Pipeline.open(result.data).inspect();
      if (!inspected.ok) throw new Error(inspected.error);
      expect(inspected.data.kids[0].kids.filter(k => k.tag === 'grid')).toHaveLength(1);
      const text = new TreeWalker().extractText(inspected.data);
      for (const token of ['제목 Alpha', '한글 표시', 'Gamma', 'Delta', '마지막 문단 Omega']) expect(text).toContain(token);
    }
  }, 90_000);
});
