import { ConversionContext } from './context';

export declare class DocxGenerator {
    private ctx;
    constructor(ctx: ConversionContext);
    convertSection(xml: string): string;
    private extractPageSetup;
    private convertParagraph;
    private getTextContent;
    private getParaPropertiesXml;
    private getRunPropertiesXml;
    private colorToHighlightName;
    private convertTable;
    private convertTableCell;
    private convertImage;
    private convertRect;
    /**
     * HWPX border type → DOCX w:val 변환
     * HWPX width(mm 문자열) → DOCX sz (1/8 포인트 단위)
     */
    private hwpBorderToDocx;
    /** tbl 레벨 <w:tblBorders> 생성 */
    private buildTblBordersXml;
    /** tc 레벨 <w:tcBorders> 생성 */
    private buildTcBordersXml;
    private escapeXml;
}
