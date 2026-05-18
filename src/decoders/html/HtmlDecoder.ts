/**
 * HtmlDecoder — HTML → DocRoot
 * 간단한 HTML 파서 (표, 이미지, 스타일 지원)
 */

import type { DocRoot, ContentNode, ParaNode, SpanNode, ImgNode, GridNode } from '../../model/doc-tree';
import type { Outcome } from '../../contract/result';
import type { TextProps, ParaProps } from '../../model/doc-props';
import { A4 } from '../../model/doc-props';
import { succeed, fail } from '../../contract/result';
import { buildRoot, buildSheet, buildPara, buildSpan, buildImg, buildGrid, buildRow, buildCell } from '../../model/builders';
import { ShieldedParser } from '../../safety/ShieldedParser';
import { TextKit } from '../../toolkit/TextKit';
import { registry } from '../../pipeline/registry';
import { BaseDecoder } from '../../core/BaseDecoder';

interface Token {
  type: 'tag' | 'text' | 'comment';
  name?: string;
  attrs?: Record<string, string>;
  selfClose?: boolean;
  close?: boolean;
  content?: string;
}

export class HtmlDecoder extends BaseDecoder {
  protected getFormat(): string { return 'html'; }

  async decode(data: Uint8Array): Promise<Outcome<DocRoot>> {
    const shield = new ShieldedParser();
    const warns: string[] = [];

    try {
      const html = this.bytesToString(data);
      const tokens = shield.guard(() => tokenize(html), [], 'html:tokenize');
      const kids = shield.guard(() => parseTokens(tokens), [], 'html:parse');

      warns.push(...shield.flush());
      const sheet = buildSheet(kids.length > 0 ? kids : [buildPara([buildSpan('')])], A4);
      return succeed(buildRoot({}, [sheet]), warns);
    } catch (e: any) {
      warns.push(...shield.flush());
      return fail(`HTML decode error: ${e?.message ?? String(e)}`, warns);
    }
  }
}

function tokenize(html: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;

  while (i < html.length) {
    if (html[i] === '<') {
      if (html[i + 1] === '!') {
        const end = html.indexOf('>', i);
        i = end + 1;
        continue;
      }

      const isClose = html[i + 1] === '/';
      const start = isClose ? i + 2 : i + 1;
      const end = html.indexOf('>', i);
      if (end === -1) break;

      const tagContent = html.slice(start, end).trim();
      const spaceIdx = tagContent.search(/\s/);
      const name = spaceIdx > 0 ? tagContent.slice(0, spaceIdx) : tagContent;
      const attrsStr = spaceIdx > 0 ? tagContent.slice(spaceIdx + 1).trim() : '';

      const attrs: Record<string, string> = {};
      if (attrsStr) {
        const attrRegex = /([a-zA-Z_:][-a-zA-Z0-9_:.]*)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g;
        let m;
        while ((m = attrRegex.exec(attrsStr)) !== null) {
          attrs[m[1].toLowerCase()] = m[2] ?? m[3] ?? m[4] ?? '';
        }
      }

      tokens.push({
        type: 'tag',
        name: name.toLowerCase(),
        attrs,
        selfClose: html[end - 1] === '/',
        close: isClose,
      });

      i = end + 1;
    } else {
      const end = html.indexOf('<', i);
      const text = end === -1 ? html.slice(i) : html.slice(i, end);
      if (text.trim()) {
        tokens.push({ type: 'text', content: text });
      }
      i = end === -1 ? html.length : end;
    }
  }

  return tokens;
}

function parseTokens(tokens: Token[]): ContentNode[] {
  const kids: ContentNode[] = [];
  let i = 0;

  while (i < tokens.length) {
    const t = tokens[i];

    if (t.type === 'tag' && !t.close) {
      switch (t.name) {
        case 'html':
          // Extract body content
          i++;
          let bodyStart = -1;
          let depth = 1;
          while (i < tokens.length && depth > 0) {
            if (tokens[i].type === 'tag' && !tokens[i].close && tokens[i].name === 'html') depth++;
            else if (tokens[i].type === 'tag' && tokens[i].close && tokens[i].name === 'html') depth--;
            else if (tokens[i].type === 'tag' && !tokens[i].close && tokens[i].name === 'body') {
              bodyStart = i + 1;
            }
            i++;
          }
          if (bodyStart > 0) {
            // Find body end
            let bodyEnd = bodyStart;
            let bodyDepth = 1;
            while (bodyEnd < tokens.length && bodyDepth > 0) {
              if (tokens[bodyEnd].type === 'tag' && !tokens[bodyEnd].close && tokens[bodyEnd].name === 'body') bodyDepth++;
              else if (tokens[bodyEnd].type === 'tag' && tokens[bodyEnd].close && tokens[bodyEnd].name === 'body') bodyDepth--;
              bodyEnd++;
            }
            bodyEnd--;
            const bodyTokens = tokens.slice(bodyStart, bodyEnd);
            const bodyKids = parseTokens(bodyTokens);
            kids.push(...bodyKids);
          }
          continue;

        case 'head':
        case 'style':
        case 'script':
          i = skipBlock(tokens, i, t.name);
          continue;

        case 'body':
        case 'div':
        case 'section':
        case 'article':
        case 'main':
          // Find closing tag and process contents
          const start = i + 1;
          let end = start;
          let divDepth = 1;
          while (end < tokens.length && divDepth > 0) {
            const t = tokens[end];
            if (t.type === 'tag' && !t.close) {
              if (['html', 'head', 'body', 'div', 'section', 'article', 'main'].includes(t.name ?? '')) divDepth++;
            } else if (t.type === 'tag' && t.close) {
              if (['html', 'head', 'body', 'div', 'section', 'article', 'main'].includes(t.name ?? '')) divDepth--;
            }
            end++;
          }
          end--;

          // Process tokens between start and end
          const subTokens = tokens.slice(start, end);
          const subKids = parseTokens(subTokens);
          kids.push(...subKids);

          i = end + 1;
          continue;

        case 'p':
          i++;
          const paraKids = collectInline(tokens, i, ['p', 'div', 'br']);
          i = paraKids.nextI;
          const align = t.attrs?.style?.includes('text-align: center') ? 'center'
                        : t.attrs?.style?.includes('text-align: right') ? 'right'
                        : t.attrs?.style?.includes('text-align: left') ? 'left'
                        : undefined;
          kids.push(buildPara(paraKids.nodes, { align }));
          continue;

        case 'br':
          kids.push(buildPara([buildSpan('')], {}));
          i++;
          continue;

        case 'img':
          i++;
          const src = t.attrs?.src;
          const alt = t.attrs?.alt || '';
          if (src?.startsWith('data:')) {
            const match = src.match(/^data:([^;]+);base64,(.+)$/);
            if (match) {
              kids.push(buildPara([buildImg(match[2], match[1] as any, 100, 100, alt)], {}));
            }
          }
          continue;

        case 'table':
          i++;
          const rows: any[] = [];
          while (i < tokens.length) {
            if (tokens[i].type === 'tag' && tokens[i].close && tokens[i].name === 'table') {
              i++;
              break;
            }
            if (tokens[i].type === 'tag' && tokens[i].name === 'tr' && !tokens[i].close) {
              i++;
              const cells: any[] = [];
              while (i < tokens.length) {
                if (tokens[i].type === 'tag' && tokens[i].close && tokens[i].name === 'tr') {
                  i++;
                  break;
                }
                if (tokens[i].type === 'tag' && (tokens[i].name === 'td' || tokens[i].name === 'th') && !tokens[i].close) {
                  i++;
                  const cellKids = collectInline(tokens, i, ['td', 'th', 'tr']);
                  i = cellKids.nextI;
                  const isHeader = tokens[i - 2]?.name === 'th';
                  const paraKids = cellKids.nodes.map(n => n.tag === 'span' ? { ...n, props: { ...n.props, b: isHeader } } : n);
                  cells.push(buildCell([buildPara(paraKids, {})]));
                } else if (tokens[i].type === 'text' && tokens[i].content?.trim()) {
                  cells.push(buildCell([buildPara([buildSpan(tokens[i].content!.trim())])]));
                  i++;
                } else {
                  i++;
                }
              }
              if (cells.length > 0) rows.push(buildRow(cells));
            } else {
              i++;
            }
          }
          if (rows.length > 0) kids.push(buildGrid(rows));
          continue;

        case 'ul':
        case 'ol':
          i++;
          const isOrdered = t.name === 'ol';
          while (i < tokens.length) {
            if (tokens[i].type === 'tag' && tokens[i].close && tokens[i].name === t.name) {
              i++;
              break;
            }
            if (tokens[i].type === 'tag' && tokens[i].name === 'li' && !tokens[i].close) {
              i++;
              const liKids = collectInline(tokens, i, ['li', 'ul', 'ol']);
              i = liKids.nextI;
              kids.push(buildPara(liKids.nodes, { listOrd: isOrdered }));
            } else {
              i++;
            }
          }
          continue;

        default:
          i++;
      }
    } else if (t.type === 'text' && t.content?.trim()) {
      kids.push(buildPara([buildSpan(t.content!.trim())], {}));
      i++;
    } else {
      i++;
    }
  }

  return kids;
}

function collectInline(tokens: Token[], start: number, stopTags: string[]): { nodes: (SpanNode | ImgNode)[]; nextI: number } {
  const nodes: (SpanNode | ImgNode)[] = [];
  let i = start;

  while (i < tokens.length) {
    const t = tokens[i];

    if (t.type === 'tag' && !t.close) {
      if (t.name && stopTags.includes(t.name)) {
        break;
      }

      switch (t.name) {
        case 'b':
        case 'strong':
          i++;
          const boldKids = collectInline(tokens, i, ['b', 'strong', ...stopTags]);
          i = boldKids.nextI;
          nodes.push(...boldKids.nodes.map(n => n.tag === 'span' ? { ...n, props: { ...n.props, b: true } } : n));
          continue;

        case 'i':
        case 'em':
          i++;
          const italicKids = collectInline(tokens, i, ['i', 'em', ...stopTags]);
          i = italicKids.nextI;
          nodes.push(...italicKids.nodes.map(n => n.tag === 'span' ? { ...n, props: { ...n.props, i: true } } : n));
          continue;

        case 'u':
          i++;
          const underlineKids = collectInline(tokens, i, ['u', ...stopTags]);
          i = underlineKids.nextI;
          nodes.push(...underlineKids.nodes.map(n => n.tag === 'span' ? { ...n, props: { ...n.props, u: true } } : n));
          continue;

        case 's':
        case 'strike':
          i++;
          const strikeKids = collectInline(tokens, i, ['s', 'strike', ...stopTags]);
          i = strikeKids.nextI;
          nodes.push(...strikeKids.nodes.map(n => n.tag === 'span' ? { ...n, props: { ...n.props, s: true } } : n));
          continue;

        case 'span':
          i++;
          const spanKids = collectInline(tokens, i, ['span', ...stopTags]);
          i = spanKids.nextI;
          const color = t.attrs?.style?.match(/color:\s*([^;]+)/)?.[1];
          nodes.push(...spanKids.nodes.map(n => n.tag === 'span' ? { ...n, props: { ...n.props, color: color || n.props.color } } : n));
          continue;

        case 'img':
          const src = t.attrs?.src;
          const alt = t.attrs?.alt || '';
          if (src?.startsWith('data:')) {
            const match = src.match(/^data:([^;]+);base64,(.+)$/);
            if (match) {
              nodes.push(buildImg(match[2], match[1] as any, 100, 100, alt));
            }
          }
          i++;
          continue;

        default:
          i++;
      }
    } else if (t.type === 'text') {
      if (t.content?.trim()) {
        nodes.push(buildSpan(t.content!.trim()));
      }
      i++;
    } else {
      i++;
    }
  }

  return { nodes: nodes.length > 0 ? nodes : [buildSpan('')], nextI: i };
}

function skipBlock(tokens: Token[], start: number, name: string): number {
  let i = start + 1;
  let depth = 1;
  while (i < tokens.length && depth > 0) {
    if (tokens[i].type === 'tag') {
      if (!tokens[i].close && tokens[i].name === name) depth++;
      if (tokens[i].close && tokens[i].name === name) depth--;
    }
    i++;
  }
  return i;
}

registry.registerDecoder(new HtmlDecoder());
