import JSZip from 'jszip';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import { Pipeline, registry, TreeWalker } from '../index';

const formats = ['hwp', 'hwpx', 'docx', 'md', 'html'];
const source = '# 제목 Alpha\n\n본문 Beta\n\n| 항목 | 값 |\n| --- | --- |\n| Gamma | Delta |';
const fixtures = new Map<string, Uint8Array>();
beforeAll(async () => {
  for (const format of formats) {
    const result = await Pipeline.open(source).to(format);
    if (!result.ok) throw new Error(result.error);
    fixtures.set(format, result.data);
  }
});

describe('public conversion matrix (content smoke test, not layout certification)', () => {
  for (const from of formats) for (const to of formats) {
    it(`${from} -> ${to} -> ${from} keeps Korean text and table cell text`, async () => {
      expect(registry.getDecoder(from)).toBeDefined();
      expect(registry.getEncoder(to)).toBeDefined();
      const converted = await Pipeline.open(fixtures.get(from)!, from).to(to);
      if (!converted.ok) throw new Error(converted.error);
      const back = await Pipeline.open(converted.data, to).to(from);
      if (!back.ok) throw new Error(back.error);
      const doc = await Pipeline.open(back.data, from).inspect();
      if (!doc.ok) throw new Error(doc.error);
      const text = new TreeWalker().extractText(doc.data);
      for (const token of ['제목', 'Alpha', '본문', 'Beta', 'Gamma', 'Delta']) expect(text).toContain(token);
    });
  }
  it.each(formats)('detects generated %s bytes without an extension', async (format) => {
    const doc = await Pipeline.open(fixtures.get(format)!).inspect();
    if (!doc.ok) throw new Error(doc.error);
    expect(new TreeWalker().extractText(doc.data)).toContain('Alpha');
  });
  it('detects a compressed DOCX with word entries beyond the first 4 KiB', async () => {
    const zip = new JSZip();
    zip.file('padding.bin', new Uint8Array(6000), { compression: 'STORE' });
    zip.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>late entry</w:t></w:r></w:p></w:body></w:document>');
    const result = await Pipeline.open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })).inspect();
    if (!result.ok) throw new Error(result.error);
    expect(new TreeWalker().extractText(result.data)).toContain('late entry');
  });
  it('registers a distinct binary DOC input and rejects fake DOC and DOC output', async () => {
    expect(registry.getDecoder('doc')).toBeDefined();
    expect(registry.getDecoder('doc')).not.toBe(registry.getDecoder('docx'));
    expect(registry.getEncoder('doc')).toBeUndefined();
    expect((await Pipeline.open('text', 'doc').to('docx')).ok).toBe(false);
    expect((await Pipeline.open('text').to('doc')).ok).toBe(false);
  });
  it('does not mistake a WordDocument OLE stream for HWP', async () => {
    const bytes = fixtures.get('hwp')!.slice();
    const name = new TextEncoder().encode('FileHeader'.split('').join('\0') + '\0');
    const offset = bytes.findIndex((_, i) => name.every((b, n) => bytes[i + n] === b));
    expect(offset).toBeGreaterThan(0);
    bytes.fill(0, offset, offset + 64);
    bytes.set(new TextEncoder().encode('WordDocument'.split('').join('\0') + '\0'), offset);
    new DataView(bytes.buffer).setUint16(offset + 64, 26, true);
    const result = await Pipeline.open(bytes).inspect();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.toLowerCase()).toContain('doc');
  });
  it('rejects an unrelated ZIP and normalizes explicit format names', async () => {
    const zip = new JSZip().file('readme.txt', 'not a document');
    expect((await Pipeline.open(await zip.generateAsync({ type: 'uint8array' })).inspect()).ok).toBe(false);
    expect((await Pipeline.open(fixtures.get('docx')!, '.DOCX').to('HWPX')).ok).toBe(true);
  });
  it('handles Blob when the runtime has no global File', async () => {
    vi.stubGlobal('File', undefined);
    try {
      const pipeline = await Pipeline.openAsync(new Blob([new Uint8Array(fixtures.get('docx')!)]));
      expect((await pipeline.inspect()).ok).toBe(true);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('detects a HWPX File whose extension incorrectly says HWP', async () => {
    const file = new File([new Uint8Array(fixtures.get('hwpx')!)], 'renamed.hwp');
    expect((await (await Pipeline.openAsync(file)).inspect()).ok).toBe(true);
    expect((await (await Pipeline.openAsync(file, 'hwp')).inspect()).ok).toBe(false);
  });

  it('does not interpret PDF or RTF bytes as Markdown', async () => {
    for (const text of ['%PDF-1.7\n', '{\\rtf1 hello}']) {
      expect((await Pipeline.open(new TextEncoder().encode(text)).to('hwp')).ok).toBe(false);
    }
  });

});
