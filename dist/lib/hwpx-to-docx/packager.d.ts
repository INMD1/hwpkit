import { ConversionContext } from './context';

export declare class DocxPackager {
    private ctx;
    constructor(ctx: ConversionContext);
    createDocxPackage(bodyXml: string): Promise<Blob>;
    private static readonly MIME_MAP;
    private isKoreanFont;
    private getDocumentFonts;
    private getDefaultFonts;
    private createContentTypes;
    private createTopRels;
    private createDocumentXml;
    private createDocumentRels;
    /** styles.xml: 문서의 실제 폰트를 기본값으로 사용 */
    private createStylesXml;
    private createSettingsXml;
    /** fontTable.xml: 문서에 사용된 폰트 전체 목록 등록 */
    private createFontTableXml;
}
