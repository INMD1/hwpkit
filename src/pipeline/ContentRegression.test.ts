import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { Pipeline, TreeWalker } from '../index';

describe('content regressions found by the conversion audit', () => {
  it('DOCX → DOCX preserves all hyperlink labels in the independent demo', async () => {
    const bytes = readFileSync('tests/fixtures/demo.docx');
    const source = await Pipeline.open(bytes).inspect();
    const result = await Pipeline.open(bytes).to('docx');
    if (!source.ok || !result.ok) throw new Error('fixture conversion failed');
    const after = await Pipeline.open(result.data).inspect();
    if (!after.ok) throw new Error(after.error);
    const text = (doc: typeof after.data) => new TreeWalker().extractText(doc).replace(/\s/g, '');
    expect(text(after.data)).toBe(text(source.data));
    const zip = await JSZip.loadAsync(result.data);
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
    expect(rels).toContain('/relationships/hyperlink');
    expect(rels).toContain('TargetMode="External"');
    expect(new TreeWalker().extractText(after.data)).toContain('calibre download page');
  });

  it('Markdown headings and table cells do not accumulate formatting markers', async () => {
    let input: Uint8Array = new TextEncoder().encode('# 제목\n\n| 항목 | 값 |\n| --- | --- |\n| **본문** | 자료 |');
    for (let i = 0; i < 3; i++) {
      const result = await Pipeline.open(input, 'md').to('md');
      if (!result.ok) throw new Error(result.error);
      input = result.data;
      const doc = await Pipeline.open(input, 'md').inspect();
      if (!doc.ok) throw new Error(doc.error);
      expect(new TreeWalker().extractText(doc.data)).toBe('제목항목값본문자료');
    }
  });

  it('HTML entities are decoded once, including numeric Unicode and inline spaces', async () => {
    const result = await Pipeline.open('<p><b>A</b> <i>B</i> &nbsp; &amp;lt; &#54620;&#xAE00;</p>', 'html').inspect();
    if (!result.ok) throw new Error(result.error);
    expect(new TreeWalker().extractText(result.data)).toBe('A B \u00a0 &lt; 한글');
    expect((await Pipeline.open('<!unfinished', 'html').inspect()).ok).toBe(true);
  });

  it('reads Markdown links and HTML table fallbacks as document content', async () => {
    const markdown = '# [**제목**](https://example.com/?a=1&b=2)\n\n' +
      '<table><tr><td>A<table><tr><td>B</td></tr></table></td><td>C</td></tr></table>';
    for (const format of ['hwp', 'hwpx', 'docx']) {
      const result = await Pipeline.open(markdown).to(format);
      if (!result.ok) throw new Error(result.error);
      const parsed = await Pipeline.open(result.data, format).inspect();
      if (!parsed.ok) throw new Error(parsed.error);
      expect(new TreeWalker().extractText(parsed.data).replace(/\s/g, '')).toBe('제목ABC');
    }
  });
});
