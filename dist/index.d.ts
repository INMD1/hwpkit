type Align = 'left' | 'center' | 'right' | 'justify';
type ImgWrap = 'inline' | 'square' | 'tight' | 'through' | 'none' | 'behind' | 'front';
type ImgHorzAlign = 'left' | 'center' | 'right';
type ImgVertAlign = 'top' | 'center' | 'bottom';
type ImgHorzRelTo = 'margin' | 'column' | 'page' | 'para';
type ImgVertRelTo = 'margin' | 'line' | 'page' | 'para';
interface ImgLayout {
    wrap: ImgWrap;
    horzAlign?: ImgHorzAlign;
    vertAlign?: ImgVertAlign;
    horzRelTo?: ImgHorzRelTo;
    vertRelTo?: ImgVertRelTo;
    xPt?: number;
    yPt?: number;
    distT?: number;
    distB?: number;
    distL?: number;
    distR?: number;
    behindDoc?: boolean;
    zOrder?: number;
}
type VAlign = 'top' | 'mid' | 'bot';
type Heading = 1 | 2 | 3 | 4 | 5 | 6;
type StrokeKind = 'solid' | 'dash' | 'dot' | 'double' | 'none';
interface TextProps {
    b?: boolean;
    i?: boolean;
    u?: boolean;
    s?: boolean;
    sup?: boolean;
    sub?: boolean;
    font?: string;
    pt?: number;
    color?: string;
    bg?: string;
}
interface ParaProps {
    align?: Align;
    heading?: Heading;
    indentPt?: number;
    spaceBefore?: number;
    spaceAfter?: number;
    lineHeight?: number;
    listLv?: number;
    listOrd?: boolean;
    listMark?: string;
}
interface Stroke {
    kind: StrokeKind;
    pt: number;
    color: string;
}
interface CellProps {
    top?: Stroke;
    bot?: Stroke;
    left?: Stroke;
    right?: Stroke;
    bg?: string;
    padPt?: number;
    align?: Align;
    va?: VAlign;
    isHeader?: boolean;
}
interface TableLook {
    firstRow?: boolean;
    lastRow?: boolean;
    firstCol?: boolean;
    lastCol?: boolean;
    bandedRows?: boolean;
    bandedCols?: boolean;
}
interface GridProps {
    widthPct?: number;
    colWidths?: number[];
    defaultStroke?: Stroke;
    look?: TableLook;
    headerRow?: boolean;
}
interface PageDims {
    wPt: number;
    hPt: number;
    mt: number;
    mb: number;
    ml: number;
    mr: number;
    orient?: 'portrait' | 'landscape';
}
interface DocMeta {
    title?: string;
    author?: string;
    subject?: string;
    desc?: string;
    keywords?: string;
    created?: string;
    modified?: string;
}
declare const A4: PageDims;
declare const A4_LANDSCAPE: PageDims;
/**
 * orient === 'landscape'일 때 wPt < hPt이면 swap,
 * orient === 'portrait'일 때 wPt > hPt이면 swap하여
 * 방향과 치수가 항상 일치하도록 정규화합니다.
 */
declare function normalizeDims(dims: PageDims): PageDims;
declare const DEFAULT_STROKE: Stroke;

type BlockTag = 'root' | 'sheet' | 'para' | 'span' | 'txt' | 'img' | 'link' | 'grid' | 'row' | 'cell' | 'br' | 'pb' | 'pagenum';
interface TxtNode {
    tag: 'txt';
    content: string;
}
interface BrNode {
    tag: 'br';
}
interface PbNode {
    tag: 'pb';
}
interface PageNumNode {
    tag: 'pagenum';
    format?: 'decimal' | 'roman' | 'romanCaps';
}
interface ImgNode {
    tag: 'img';
    b64: string;
    mime: 'image/png' | 'image/jpeg' | 'image/gif' | 'image/bmp';
    w: number;
    h: number;
    alt?: string;
    layout?: ImgLayout;
}
interface SpanNode {
    tag: 'span';
    props: TextProps;
    kids: (TxtNode | BrNode | PbNode | PageNumNode)[];
}
interface LinkNode {
    tag: 'link';
    href: string;
    kids: SpanNode[];
}
interface ParaNode {
    tag: 'para';
    props: ParaProps;
    kids: (SpanNode | ImgNode | LinkNode)[];
}
interface CellNode {
    tag: 'cell';
    cs: number;
    rs: number;
    props: CellProps;
    kids: ParaNode[];
}
interface RowNode {
    tag: 'row';
    kids: CellNode[];
    heightPt?: number;
}
interface GridNode {
    tag: 'grid';
    props: GridProps;
    kids: RowNode[];
}
type ContentNode = ParaNode | GridNode;
interface SheetNode {
    tag: 'sheet';
    dims: PageDims;
    kids: ContentNode[];
    header?: ParaNode[];
    footer?: ParaNode[];
}
interface DocRoot {
    tag: 'root';
    meta: DocMeta;
    kids: SheetNode[];
}
type AnyNode = DocRoot | SheetNode | ParaNode | SpanNode | TxtNode | ImgNode | LinkNode | GridNode | RowNode | CellNode | BrNode | PbNode | PageNumNode;

type Outcome<T> = Ok<T> | Fail;
interface Ok<T> {
    ok: true;
    data: T;
    warns: string[];
}
interface Fail {
    ok: false;
    error: string;
    warns: string[];
}
declare function succeed<T>(data: T, warns?: string[]): Ok<T>;
declare function fail(error: string, warns?: string[]): Fail;

interface Decoder {
    readonly format: string;
    decode(data: Uint8Array): Promise<Outcome<DocRoot>>;
}

interface Encoder {
    readonly format: string;
    encode(doc: DocRoot): Promise<Outcome<Uint8Array>>;
}

declare class Pipeline {
    private raw;
    private srcFmt;
    private constructor();
    /** 파일을 열고 포맷을 자동 감지하거나 명시 */
    static open(input: Uint8Array | string, fmt?: string): Pipeline;
    /** File/Blob 비동기 입력 */
    static openAsync(input: File | Blob | Uint8Array | string, fmt?: string): Promise<Pipeline>;
    /** 목표 포맷으로 변환 */
    to(targetFmt: string): Promise<Outcome<Uint8Array>>;
    /** DocRoot만 추출 (인코딩 없이) */
    inspect(): Promise<Outcome<DocRoot>>;
}

declare class FormatRegistry {
    private decoders;
    private encoders;
    registerDecoder(d: Decoder): void;
    registerEncoder(e: Encoder): void;
    getDecoder(fmt: string): Decoder | undefined;
    getEncoder(fmt: string): Encoder | undefined;
    supportedInputs(): string[];
    supportedOutputs(): string[];
}
declare const registry: FormatRegistry;

declare function buildRoot(meta?: DocMeta, kids?: SheetNode[]): DocRoot;
declare function buildSheet(kids?: ContentNode[], dims?: PageDims, opts?: {
    header?: ParaNode[];
    footer?: ParaNode[];
}): SheetNode;
declare function buildPageNum(format?: PageNumNode['format']): PageNumNode;
declare function buildBr(): BrNode;
declare function buildPb(): PbNode;
declare function buildPara(kids?: ParaNode['kids'], props?: ParaProps): ParaNode;
declare function buildSpan(content: string, props?: TextProps): SpanNode;
declare function buildImg(b64: string, mime: ImgNode['mime'], w: number, h: number, alt?: string, layout?: ImgLayout): ImgNode;
declare function buildGrid(kids: RowNode[], props?: GridProps): GridNode;
declare function buildRow(kids: CellNode[], heightPt?: number): RowNode;
declare function buildCell(kids: ParaNode[], opts?: {
    cs?: number;
    rs?: number;
    props?: CellProps;
}): CellNode;

declare class ShieldedParser {
    private log;
    /** 단일 요소 안전 파싱 */
    guard<T>(fn: () => T, fallback: T, label: string): T;
    /** 배열 각 요소 독립 파싱 (하나 실패해도 나머지 계속) */
    guardAll<I, O>(items: I[], fn: (x: I, i: number) => O, fb: (x: I, i: number) => O, label: string): O[];
    /**
     * 표 전용 4단계 폴백
     *   Lv1: Full → Lv2: Grid → Lv3: Flat → Lv4: Text
     */
    guardGrid<T>(node: unknown, lv1Full: (n: unknown) => T, lv2Grid: (n: unknown) => T, lv3Flat: (n: unknown) => T, lv4Text: (n: unknown) => T, label: string): {
        value: T;
        level: 1 | 2 | 3 | 4;
    };
    /** 이미지 안전 파싱 */
    guardImg<T>(node: unknown, fn: (n: unknown) => T, placeholder: (alt: string) => T, label: string): T;
    private warn;
    flush(): string[];
}

declare const Metric: {
    readonly hwpToPt: (v: number) => number;
    readonly ptToHwp: (v: number) => number;
    readonly hwpToDxa: (v: number) => number;
    readonly dxaToHwp: (v: number) => number;
    readonly hwpToEmu: (v: number) => number;
    readonly emuToHwp: (v: number) => number;
    readonly dxaToPt: (v: number) => number;
    readonly ptToDxa: (v: number) => number;
    readonly dxaToEmu: (v: number) => number;
    readonly emuToDxa: (v: number) => number;
    readonly emuToPt: (v: number) => number;
    readonly ptToEmu: (v: number) => number;
    readonly hHeightToPt: (v: number) => number;
    readonly ptToHHeight: (v: number) => number;
    readonly halfPtToPt: (v: number) => number;
    readonly ptToHalfPt: (v: number) => number;
};
declare function safeHex(raw: string | number | null | undefined): string | undefined;
declare function safeAlign(raw?: string): Align;
declare function safeStrokeHwpx(type?: string, w?: number, c?: string): Stroke;
declare function safeStrokeDocx(val?: string, sz?: number, c?: string): Stroke;
declare function safeFont(raw?: string): string;
declare function safeFontToKr(raw?: string): string;

type WalkCallback = (node: AnyNode, parent: AnyNode | null, depth: number) => void | 'stop';
declare function walkNode(node: AnyNode, cb: WalkCallback, parent?: AnyNode | null, depth?: number): boolean;
declare class TreeWalker {
    walk(root: DocRoot, cb: WalkCallback): void;
    findAll<T extends AnyNode>(root: DocRoot, predicate: (n: AnyNode) => n is T): T[];
    extractText(root: DocRoot): string;
}

declare function countNodes(root: DocRoot): Record<string, number>;
declare function validateRoot(root: DocRoot): string[];

declare const XmlKit: {
    /** @deprecated Use parseStrict instead */
    parse(xml: string): Promise<unknown>;
    parseStrict(xml: string): Promise<unknown>;
    attr(node: Record<string, unknown>, key: string): string | undefined;
    text(node: Record<string, unknown> | string | undefined): string;
};

interface ZipEntry {
    name: string;
    data: Uint8Array;
}
declare const ArchiveKit: {
    inflate(compressed: Uint8Array): Promise<Uint8Array>;
    deflate(data: Uint8Array): Promise<Uint8Array>;
    unzip(zipData: Uint8Array): Promise<Map<string, Uint8Array>>;
    zip(entries: ZipEntry[]): Promise<Uint8Array>;
};

/**
 * OLE2 Compound File Binary Format (CFB) parser.
 * Used for legacy HWP 5.0 files.
 */
declare const BinaryKit: {
    readU16LE(buf: Uint8Array, offset: number): number;
    readU32LE(buf: Uint8Array, offset: number): number;
    isOle2(data: Uint8Array): boolean;
    parseCfb(data: Uint8Array): Map<string, Uint8Array>;
};

declare const TextKit: {
    decode(data: Uint8Array, encoding?: string): string;
    encode(text: string): Uint8Array;
    escapeXml(s: string): string;
    unescapeXml(s: string): string;
    normalizeWhitespace(s: string): string;
    stripControl(s: string): string;
    base64Encode(data: Uint8Array): string;
    base64Decode(b64: string): Uint8Array;
};

export { A4, A4_LANDSCAPE, type Align, type AnyNode, ArchiveKit, BinaryKit, type BlockTag, type BrNode, type CellNode, type CellProps, type ContentNode, DEFAULT_STROKE, type Decoder, type DocMeta, type DocRoot, type Encoder, type Fail, type GridNode, type GridProps, type Heading, type ImgNode, type LinkNode, Metric, type Ok, type Outcome, type PageDims, type PageNumNode, type ParaNode, type ParaProps, type PbNode, Pipeline, type RowNode, type SheetNode, ShieldedParser, type SpanNode, type Stroke, type StrokeKind, type TableLook, TextKit, type TextProps, TreeWalker, type TxtNode, type VAlign, XmlKit, buildBr, buildCell, buildGrid, buildImg, buildPageNum, buildPara, buildPb, buildRoot, buildRow, buildSheet, buildSpan, countNodes, fail, normalizeDims, registry, safeAlign, safeFont, safeFontToKr, safeHex, safeStrokeDocx, safeStrokeHwpx, succeed, validateRoot, walkNode };
