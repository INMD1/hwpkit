import { ConversionState } from './state';

export declare class HwpxGenerator {
    private state;
    constructor(state: ConversionState);
    private createFontFaces;
    private createSection;
    private createSecPr;
    private convertParagraph;
    private convertRun;
    private convertFootnote;
    private getParaShapeId;
    private getCharShapeId;
    private getOrCreateHighlightBorderFill;
    private registerFontsFromRFonts;
    private getParagraphFontSize;
    private getListPrefix;
    private getCellBorderFill;
    private convertTable;
    private convertDrawing;
    private convertPicture;
    private createParaProperties;
    private createCharProperties;
    private createContentHpf;
    private createContainer;
    private createContainerhdf;
    private createManifest;
    createHwpx(doc: Document): any;
    private createVersion;
    private createSettings;
    private createHeader;
    private createBorderFills;
    private createBinDataProperties;
    private createTabProperties;
    private createNumberings;
    private convertLvlTextToHwpx;
    private createStyles;
}
