import { MdConversionContext } from './context';

export declare class HwpxMdParser {
    private ctx;
    constructor(ctx: MdConversionContext);
    parseHeader(xml: string): void;
    private parseCharProperties;
    private parseParaProperties;
    private parseStyles;
    private parseBorderFills;
}
