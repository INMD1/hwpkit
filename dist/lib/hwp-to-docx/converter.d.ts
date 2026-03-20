/**
 * HWP → DOCX 변환 오케스트레이터
 * HwpParser로 HWP 바이너리를 파싱하고,
 * DocxGenerator로 DOCX XML을 생성한 뒤,
 * 최종 DOCX 패키지(ZIP)를 만듭니다.
 */
export declare class HwpToDocxConverter {
    convert(input: File | ArrayBuffer | Blob): Promise<Blob>;
    private createDocxPackage;
    private static readonly MIME_MAP;
    private isKoreanFont;
    private getDefaultFonts;
    private contentTypesXml;
    private topRelsXml;
    private documentXml;
    private documentRelsXml;
    /** styles.xml: 문서의 실제 폰트를 기본값으로 사용 */
    private stylesXml;
    /** fontTable.xml: 문서에 사용된 폰트 전체 목록 등록 */
    private fontTableXml;
}
