import { MdConversionContext } from './context';

export declare class HwpxMdGenerator {
    private ctx;
    constructor(ctx: MdConversionContext);
    convertSection(xml: string): string;
    private convertParagraph;
    private applyInlineFormatting;
    private getTextContent;
    private expandCells;
    private convertTable;
    private getCellSizeStyle;
    private getCellBorderStyle;
    private extractCellText;
    private convertImage;
    private uint8ArrayToBase64;
    private convertRect;
}
