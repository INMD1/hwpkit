import { SaxesParser } from 'saxes';

// Parses XML into an xml2js-compatible object structure:
//   { tagName: [{ _attr: { k: v }, _text: '...', childTag: [...] }] }
function parseXmlStrict(xml: string): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const parser = new SaxesParser({ xmlns: false });

    interface Frame { tag: string; obj: Record<string, unknown> }
    const stack: Frame[] = [];
    let result: unknown = null;

    parser.on('error', (err: Error) => reject(err));

    parser.on('opentag', (node: any) => {
      const obj: Record<string, unknown> = {};
      const attrs = node.attributes as Record<string, string>;
      if (attrs && Object.keys(attrs).length > 0) {
        obj['_attr'] = { ...attrs };
      }
      stack.push({ tag: node.name as string, obj });
    });

    const appendText = (text: string) => {
      if (stack.length > 0 && text) {
        const frame = stack[stack.length - 1];
        const cur = frame.obj['_text'];
        frame.obj['_text'] = typeof cur === 'string' ? cur + text : text;
      }
    };

    parser.on('text', (text: string) => appendText(text));
    parser.on('cdata', (cdata: string) => appendText(cdata));

    parser.on('closetag', () => {
      const frame = stack.pop();
      if (!frame) return;
      const { tag, obj } = frame;

      // Drop whitespace-only _text
      if (typeof obj['_text'] === 'string' && !(obj['_text'] as string).trim()) {
        delete obj['_text'];
      }

      if (stack.length === 0) {
        result = { [tag]: [obj] };
      } else {
        const parent = stack[stack.length - 1].obj;
        const existing = parent[tag];
        if (Array.isArray(existing)) {
          (existing as unknown[]).push(obj);
        } else {
          parent[tag] = [obj];
        }
        // Track child order for document-order iteration
        if (!parent['_childOrder']) parent['_childOrder'] = [];
        (parent['_childOrder'] as string[]).push(tag);
      }
    });

    try {
      parser.write(xml).close();
      resolve(result);
    } catch (e) {
      reject(e);
    }
  });
}

export const XmlKit = {
  /** @deprecated Use parseStrict instead */
  async parse(xml: string): Promise<unknown> {
    return parseXmlStrict(xml);
  },

  async parseStrict(xml: string): Promise<unknown> {
    return parseXmlStrict(xml);
  },

  attr(node: Record<string, unknown>, key: string): string | undefined {
    const a = node['_attr'] as Record<string, string> | undefined;
    return a?.[key];
  },

  text(node: Record<string, unknown> | string | undefined): string {
    if (node == null) return '';
    if (typeof node === 'string') return node;
    const t = node['_text'];
    return typeof t === 'string' ? t : '';
  },
};
