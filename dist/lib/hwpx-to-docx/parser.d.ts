import { ConversionContext } from './context';

export declare class HwpxParser {
    private ctx;
    constructor(ctx: ConversionContext);
    parseContentHpf(_xml: string): Promise<void>;
    parseHeader(xml: string): void;
    private parseFontFaces;
    private parseBorderFills;
    private parseCharProperties;
    private parseParaProperties;
    private parseStyles;
    private parseNumberings;
}
