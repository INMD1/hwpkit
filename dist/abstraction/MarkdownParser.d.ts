export interface PandocNode {
    t: string;
    c?: any;
}
export interface PandocAst {
    'pandoc-api-version': number[];
    meta: any;
    blocks: PandocNode[];
}
export interface IMarkdownParser {
    parse(markdown: string): Promise<PandocAst>;
}
