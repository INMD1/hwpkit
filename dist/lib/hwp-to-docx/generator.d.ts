import { CharShape, ParaShape, BorderFill, TableCell, PageDef, TextSegment } from './type';

/**
 * DOCX XML 생성 담당
 * HWP에서 파싱된 데이터를 기반으로 DOCX XML 요소를 직접 생성합니다.
 * paragraph, run, table, image 등의 XML을 담당합니다.
 */
export declare class DocxGenerator {
    private faceNames;
    private charShapes;
    private paraShapes;
    private pageDefs;
    private borderFills;
    private nextImageId;
    constructor(faceNames: string[], charShapes: CharShape[], paraShapes: ParaShape[], pageDefs: PageDef[], borderFills?: BorderFill[]);
    /** 섹션 하나의 바이너리 스트림을 파싱하여 DOCX body/header/footer XML을 반환 */
    parseSection(stream: Uint8Array): {
        body: string;
        header: string | null;
        footer: string | null;
    };
    /** PARA_TEXT + PARA_CHAR_SHAPE + PARA_SHAPE + LINE_SEG → 완전한 <w:p> */
    buildParagraphXml(textData: Uint8Array, charShapeMap: {
        pos: number;
        shapeId: number;
    }[], paraShapeId: number, inTableCell?: boolean, lineSegs?: TextSegment[]): string;
    /** ParaShape → <w:pPr> */
    buildParaPropsXml(paraShapeId: number, inTableCell?: boolean, pageBreakBefore?: boolean): string;
    /** PARA_TEXT 바이너리에서 <w:r> 요소들 직접 생성
     * @param inlineBreaks  LINE_SEG에서 추출한 {문자위치 → 'page'|'column'} 맵.
     *                      해당 위치 직전에 <w:br w:type="..."/> 를 삽입한다.
     */
    buildRunsFromText(textData: Uint8Array, charShapeMap: {
        pos: number;
        shapeId: number;
    }[], inlineBreaks?: Map<number, 'page' | 'column'>): string;
    /** 텍스트 + CharShape ID → <w:r> XML (서식 포함) */
    buildRunXml(text: string, shapeId: number): string;
    /** tableData 2D 배열 → 완전한 <w:tbl> XML */
    buildTableXml(tableData: (TableCell | null)[][], alignment?: number, tableBorderFillId?: number): string;
    /** 2D 배열에 셀 데이터 채우기 (병합 처리 포함) */
    fillTableData(tableData: (TableCell | null)[][], rIdx: number, cIdx: number, cell: TableCell): void;
    /** DOCX DrawingML 이미지 요소 생성 */
    buildImageXml(binDataId: number, widthTwips: number, heightTwips: number): string;
    /** 구역 설정 <w:sectPr> XML */
    buildSectPrXml(hasHeader?: boolean, hasFooter?: boolean): string;
    escapeXml(text: string): string;
    private alignToDocx;
    private hwpBorderTypeToDocx;
    /** HWP 선 굵기 코드 → DOCX w:sz (1/8pt) */
    private hwpThicknessToDocxSz;
    /** BorderFill의 한 방향 → DOCX <w:${side}> XML 문자열 */
    private buildBorderXml;
}
