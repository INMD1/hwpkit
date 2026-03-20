export declare class HwpParser {
    private ole;
    private charShapes;
    private paraShapes;
    private borderFills;
    private compressed;
    constructor(data: Uint8Array);
    parse(): string;
    private parseFileHeader;
    private parseDocInfo;
    private parseBodyText;
    private buildCellStyle;
}
