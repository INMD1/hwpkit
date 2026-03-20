/**
 * React Hooks - HWPKit
 *
 * 모든 Hook은 useTransform 제네릭 훅을 기반으로 합니다.
 * 새로운 Transformer API (MdTransformer.fromHwpx 등)를 React 컴포넌트에서 쉽게 사용할 수 있습니다.
 */

import { useState, useCallback } from 'react';
import { MdTransformer } from '../transformers/MdTransformer';
import { HwpxTransformer } from '../transformers/HwpxTransformer';
import { DocxTransformer } from '../transformers/DocxTransformer';

// ─── 공통 훅 상태 ─────────────────────────────────────────────────────────────

export interface TransformState<T> {
  result: T | null;
  isPending: boolean;
  error: Error | null;
}

// ─── 제네릭 기반 훅 ───────────────────────────────────────────────────────────

export function useTransform<TInput, TOutput>(
  transformFn: (input: TInput) => Promise<TOutput>,
): TransformState<TOutput> & { run: (input: TInput) => Promise<TOutput | null> } {
  const [state, setState] = useState<TransformState<TOutput>>({
    result: null,
    isPending: false,
    error: null,
  });

  const run = useCallback(async (input: TInput): Promise<TOutput | null> => {
    setState({ result: null, isPending: true, error: null });
    try {
      const result = await transformFn(input);
      setState({ result, isPending: false, error: null });
      return result;
    } catch (e) {
      const error = e instanceof Error ? e : new Error(String(e));
      setState({ result: null, isPending: false, error });
      return null;
    }
  }, [transformFn]);

  return { ...state, run };
}

// ─── 구체적인 훅들 ────────────────────────────────────────────────────────────

/** HWPX → Markdown */
export function useHwpxToMd() {
  return useTransform<File | Blob | Uint8Array, string>(MdTransformer.fromHwpx);
}

/** HWP → Markdown */
export function useHwpToMd() {
  return useTransform<File | Blob | Uint8Array, string>(MdTransformer.fromHwp);
}

/** DOCX → Markdown */
export function useDocxToMd() {
  return useTransform<File | Blob | Uint8Array, string>(MdTransformer.fromDocx);
}

/** DOCX → HWPX */
export function useDocxToHwpx() {
  return useTransform<File | Blob | Uint8Array, Blob>(HwpxTransformer.fromDocx);
}

/** Markdown → HWPX */
export function useMdToHwpx() {
  return useTransform<string, Blob>(HwpxTransformer.fromMd);
}

/** HWPX → DOCX */
export function useHwpxToDocx() {
  return useTransform<File | Blob | Uint8Array, Blob>(DocxTransformer.fromHwpx);
}

/** HWP → DOCX */
export function useHwpToDocx() {
  return useTransform<File | Blob | Uint8Array, Blob>(DocxTransformer.fromHwp);
}

/** Markdown → DOCX */
export function useMdToDocx() {
  return useTransform<string, Blob>(DocxTransformer.fromMd);
}

// ─── 다운로드 헬퍼 ────────────────────────────────────────────────────────────

export function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}
