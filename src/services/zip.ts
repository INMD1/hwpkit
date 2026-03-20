/**
 * ZipAdapter - fflate 기반 ZIP 처리 서비스
 *
 * 기존 JSZip 대신 fflate를 사용합니다.
 * fflate는 순수 TypeScript로 작성된 경량 압축 라이브러리로
 * 브라우저/Node.js 양쪽에서 동작합니다.
 */

import { unzip, zip, type Unzipped, type AsyncUnzipOptions } from 'fflate';

export type ZipEntries = Record<string, Uint8Array>;

// ─── Read ─────────────────────────────────────────────────────────────────────

export function extractZip(data: Uint8Array): Promise<ZipEntries> {
  return new Promise((resolve, reject) => {
    const opts: AsyncUnzipOptions = {};
    unzip(data, opts, (err, files) => {
      if (err) reject(new Error(`[ZipAdapter] 압축 해제 실패: ${err.message}`));
      else resolve(files as ZipEntries);
    });
  });
}

export function getEntry(entries: ZipEntries, path: string): Uint8Array | undefined {
  // 경로 구분자 정규화 (Windows 스타일 → Unix)
  const normalized = path.replace(/\\/g, '/');
  return entries[normalized] ?? entries[normalized.replace(/^\//, '')];
}

export function requireEntry(entries: ZipEntries, path: string): Uint8Array {
  const data = getEntry(entries, path);
  if (!data) throw new Error(`[ZipAdapter] 엔트리 없음: ${path}`);
  return data;
}

export function entryToText(data: Uint8Array): string {
  return new TextDecoder('utf-8').decode(data);
}

export function listEntries(entries: ZipEntries, prefix: string): string[] {
  const norm = prefix.replace(/\\/g, '/').replace(/\/$/, '');
  return Object.keys(entries).filter(k => k.startsWith(norm + '/') || k.startsWith(norm));
}

// ─── Write ────────────────────────────────────────────────────────────────────

export interface ZipWriteEntry {
  path: string;
  data: Uint8Array | string;
}

export function buildZip(entries: ZipWriteEntry[]): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const input: Record<string, Uint8Array> = {};
    for (const { path, data } of entries) {
      input[path] = typeof data === 'string'
        ? new TextEncoder().encode(data)
        : data;
    }
    zip(input, (err, result) => {
      if (err) reject(new Error(`[ZipAdapter] 압축 생성 실패: ${err.message}`));
      else resolve(result);
    });
  });
}

export function zipToBlob(data: Uint8Array): Blob {
  // Uint8Array를 ArrayBuffer로 변환하여 Blob 생성
  const buf = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
  return new Blob([buf], { type: 'application/zip' });
}
