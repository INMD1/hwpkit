export const TextKit = {
  decode(data: Uint8Array, encoding = 'utf-8'): string {
    try {
      return new TextDecoder(encoding, { fatal: true }).decode(data);
    } catch {
      return new TextDecoder('utf-8', { fatal: false }).decode(data);
    }
  },

  encode(text: string): Uint8Array {
    return new TextEncoder().encode(text);
  },

  escapeXml(s: string): string {
    return s
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  },

  unescapeXml(s: string): string {
    return s
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
  },

  normalizeWhitespace(s: string): string {
    return s.replace(/\s+/g, ' ').trim();
  },

  stripControl(s: string): string {
    // eslint-disable-next-line no-control-regex
    return s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
  },

  base64Encode(data: Uint8Array): string {
    let binary = '';
    for (let i = 0; i < data.length; i++) {
      binary += String.fromCharCode(data[i]);
    }
    return btoa(binary);
  },

  base64Decode(b64: string): Uint8Array {
    const binary = atob(b64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  },
};
