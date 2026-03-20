import { CharShape, ParaShape, BorderFill, PageDef } from './type';
import * as cfb from 'cfb';
/**
 * HWP 바이너리 파싱 담당
 * OLE (CFB) 구조에서 FileHeader, DocInfo, BodyText를 읽어
 * 글꼴 이름, 글자 모양, 문단 모양, 페이지 설정 등을 추출합니다.
 */
/** 이미지 관계 타입 */
export interface ImageRel {
    id: string;
    target: string;
    data: Uint8Array;
    ext: string;
}
/** 파싱 결과 타입 */
export interface HwpParseResult {
    faceNames: string[];
    charShapes: CharShape[];
    paraShapes: ParaShape[];
    pageDefs: PageDef[];
    borderFills: BorderFill[];
    imageRels: ImageRel[];
    compressed: boolean;
    ole: cfb.CFB$Container;
}
export declare class HwpParser {
    private ole;
    private compressed;
    faceNames: string[];
    charShapes: CharShape[];
    paraShapes: ParaShape[];
    pageDefs: PageDef[];
    borderFills: BorderFill[];
    imageRels: ImageRel[];
    /**
     * ArrayBuffer에서 HWP OLE를 읽고 DocInfo를 파싱합니다.
     * @returns 파싱된 OLE 컨테이너와 압축 여부
     */
    parse(data: Uint8Array): HwpParseResult;
    /** OLE에서 BodyText 섹션 스트림 목록 반환 */
    getSectionStreams(): Uint8Array[];
    private parseFileHeader;
    private parseDocInfo;
    /** HWPTAG_FACE_NAME(17) - 글꼴 이름 추출 */
    private parseFaceNameRecord;
    /**
     * HWPTAG_CHAR_SHAPE(21) - 글자 모양 파싱
     * 오프셋 레이아웃:
     *   0~13:  WORD[7]  언어별 글꼴 ID (14바이트)
     *   14~20: UINT8[7] 장평 비율
     *   21~27: INT8[7]  자간
     *   28~34: UINT8[7] 상대크기
     *   35~41: INT8[7]  글자 위치
     *   42~45: INT32    기준 크기 (HWPUNIT, /100 → pt)
     *   46~49: UINT32   속성 (굵기, 기울임, 밑줄, 취소선 등)
     *   50~51: INT8×2   그림자 간격
     *   52~55: COLORREF 글자 색
     */
    private parseCharShapeRecord;
    /**
     * HWPTAG_PARA_SHAPE(25) - 문단 모양 파싱
     * 레이아웃:
     *   +0  UINT32 속성1 (bit0-1: 줄간격종류, bit2-4: 정렬)
     *   +4  INT32  왼쪽 여백
     *   +8  INT32  오른쪽 여백
     *   +12 INT32  들여쓰기/내어쓰기
     *   +16 INT32  문단 간격 위
     *   +20 INT32  문단 간격 아래
     *   +24 INT32  줄 간격 (5.0.2.5 미만)
     *   +28 UINT16 탭 정의 ID
     *   +30 UINT16 번호/글머리표 ID
     *   +32 UINT16 테두리/배경 ID
     *   +34 INT16  테두리 좌 간격
     *   +36 INT16  테두리 우 간격
     *   +38 INT16  테두리 상 간격
     *   +40 INT16  테두리 하 간격
     *   +42 UINT32 속성2 (5.0.1.7+)
     *   +46 UINT32 속성3 (5.0.2.5+)
     *   +50 UINT32 줄 간격 (5.0.2.5+)
     */
    private parseParaShapeRecord;
    /** HWPTAG_BORDER_FILL(20) - 테두리/배경 */
    private parseBorderFillRecord;
    /** OLE BinData 스토리지에서 이미지 파일 추출 */
    private extractImages;
}
