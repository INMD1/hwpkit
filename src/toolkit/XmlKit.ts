import { parseStringPromise, Builder } from 'xml2js';

export const XmlKit = {
  async parse(xml: string): Promise<unknown> {
    return parseStringPromise(xml, {
      explicitArray: false,
      mergeAttrs: false,
      explicitCharkey: true,
      charkey: '_text',
      attrkey: '_attr',
    });
  },

  async parseStrict(xml: string): Promise<unknown> {
    return parseStringPromise(xml, {
      explicitArray: true,
      mergeAttrs: false,
      explicitCharkey: true,
      charkey: '_text',
      attrkey: '_attr',
    });
  },

  build(obj: unknown, opts: { rootName?: string; headless?: boolean } = {}): string {
    const builder = new Builder({
      rootName: opts.rootName ?? 'root',
      headless: opts.headless ?? false,
      renderOpts: { pretty: false },
    });
    return builder.buildObject(obj);
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
