/**
 * EncodingAdapter - 문자 인코딩 변환 서비스
 * iconv-lite를 래핑합니다.
 */

import * as iconv from 'iconv-lite';

export function decodeBuffer(
  data: Uint8Array | Buffer,
  encoding: string,
): string {
  return iconv.decode(Buffer.from(data), encoding);
}

export function encodeString(text: string, encoding: string): Uint8Array {
  const buf = iconv.encode(text, encoding);
  return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
}

export function toBase64(data: Uint8Array): string {
  let binary = '';
  for (let i = 0; i < data.length; i++) binary += String.fromCharCode(data[i]);
  return btoa(binary);
}

export function fromBase64(b64: string): Uint8Array {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

export function detectMimeType(data: Uint8Array): 'image/png' | 'image/jpeg' | 'image/gif' | 'image/webp' {
  if (data[0] === 0x89 && data[1] === 0x50) return 'image/png';
  if (data[0] === 0xFF && data[1] === 0xD8) return 'image/jpeg';
  if (data[0] === 0x47 && data[1] === 0x49) return 'image/gif';
  if (data[8] === 0x57 && data[9] === 0x45) return 'image/webp';
  return 'image/png';
}
