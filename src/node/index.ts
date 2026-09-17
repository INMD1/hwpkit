import { mkdtemp, readFile, writeFile, mkdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { spawn } from 'node:child_process';

export interface LibreOfficeDocOptions {
  /** Executable path; arguments are passed directly, without a shell. */
  executable?: string;
  timeoutMs?: number;
}

/** Local DOC → DOCX bridge. Requires LibreOffice, never uploads documents. */
export function createLibreOfficeDocConverter(options: LibreOfficeDocOptions = {}): (data: Uint8Array) => Promise<Uint8Array> {
  const executable = options.executable ?? 'libreoffice';
  const timeoutMs = options.timeoutMs ?? 60_000;
  if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) throw new Error('timeoutMs must be positive');
  return async data => {
    const directory = await mkdtemp(join(tmpdir(), 'hwpkit-doc-'));
    try {
      const input = join(directory, 'input.doc');
      const output = join(directory, 'output');
      await writeFile(input, data);
      await mkdir(output);
      await new Promise<void>((resolve, reject) => {
        const child = spawn(executable, [
          `-env:UserInstallation=${pathToFileURL(join(directory, 'profile')).href}`,
          '--headless', '--convert-to', 'docx:Office Open XML Text', '--outdir', output, input,
        ], { stdio: ['ignore', 'pipe', 'pipe'], detached: process.platform !== 'win32' });
        let detail = '';
        let timedOut = false;
        const capture = (chunk: Buffer) => { detail = (detail + chunk.toString()).slice(-4000); };
        child.stdout.on('data', capture);
        child.stderr.on('data', capture);
        const timer = setTimeout(() => {
          timedOut = true;
          // Kill the process group as soffice can launch a child process.
          try {
            if (process.platform !== 'win32' && child.pid) process.kill(-child.pid, 'SIGKILL');
            else child.kill('SIGKILL');
          } catch { child.kill('SIGKILL'); }
        }, timeoutMs);
        child.on('error', error => {
          clearTimeout(timer);
          reject(new Error(`LibreOffice 실행 실패 (${executable}): ${error.message}`));
        });
        child.on('close', code => {
          clearTimeout(timer);
          if (timedOut) reject(new Error(`DOC 변환 시간 초과 (${timeoutMs}ms)`));
          else if (code !== 0) reject(new Error(`LibreOffice DOC 변환 실패 (${code}): ${detail.trim()}`));
          else resolve();
        });
      });
      // LibreOffice can exit 0 even if it failed to load the document.
      let result: Uint8Array;
      try { result = new Uint8Array(await readFile(join(output, 'input.docx'))); }
      catch { throw new Error('LibreOffice가 DOCX를 생성하지 못했습니다. DOC 파일을 확인하세요.'); }
      if (result[0] !== 0x50 || result[1] !== 0x4b) throw new Error('LibreOffice 출력이 DOCX가 아닙니다.');
      return result;
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  };
}
