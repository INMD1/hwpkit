import * as cfb from 'cfb';
import pako from 'pako';
import { CharShape, ParaShape, BorderFill, PageDef } from './type';

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

export class HwpParser {
    private ole!: cfb.CFB$Container;
    private compressed: boolean = false;

    // 파싱된 정보 저장
    faceNames: string[] = [];
    charShapes: CharShape[] = [];
    paraShapes: ParaShape[] = [];
    pageDefs: PageDef[] = [];
    borderFills: BorderFill[] = [];
    imageRels: ImageRel[] = [];

    /**
     * ArrayBuffer에서 HWP OLE를 읽고 DocInfo를 파싱합니다.
     * @returns 파싱된 OLE 컨테이너와 압축 여부
     */
    parse(data: Uint8Array): HwpParseResult {
        this.ole = cfb.read(data, { type: 'buffer' });
        this.parseFileHeader();
        this.parseDocInfo();
        this.extractImages();

        return {
            faceNames: this.faceNames,
            charShapes: this.charShapes,
            paraShapes: this.paraShapes,
            pageDefs: this.pageDefs,
            borderFills: this.borderFills,
            imageRels: this.imageRels,
            compressed: this.compressed,
            ole: this.ole,
        };
    }

    /** OLE에서 BodyText 섹션 스트림 목록 반환 */
    getSectionStreams(): Uint8Array[] {
        const sections = this.ole.FullPaths
            .filter(p => p.toLowerCase().includes('bodytext/section'))
            .sort((a, b) => {
                const aMatch = a.toLowerCase().match(/section(\d+)/);
                const bMatch = b.toLowerCase().match(/section(\d+)/);
                const aNum = aMatch ? parseInt(aMatch[1], 10) : 0;
                const bNum = bMatch ? parseInt(bMatch[1], 10) : 0;
                return aNum - bNum;
            });

        if (sections.length === 0) throw new Error('BodyText sections not found');

        const streams: Uint8Array[] = [];
        for (const secPath of sections) {
            const secEntry = cfb.find(this.ole, secPath);
            if (!secEntry || !secEntry.content) continue;

            let stream = new Uint8Array(secEntry.content as any);
            if (this.compressed) {
                try { stream = pako.inflate(stream, { raw: true }); } catch (_) { continue; }
            }
            streams.push(stream);
        }
        return streams;
    }

    // ─── FileHeader ──────────────────────────────────────────

    private parseFileHeader() {
        const fileHeaderEntry = this.ole.FullPaths.find(p => p.toLowerCase().endsWith('fileheader'));
        if (!fileHeaderEntry) throw new Error('Not a valid HWP 5.0 file');

        const fileHeader = cfb.find(this.ole, fileHeaderEntry);
        if (!fileHeader || !fileHeader.content || fileHeader.content.length < 40)
            throw new Error('Invalid FileHeader');

        const headerContent = new Uint8Array(fileHeader.content as any);
        const headerView = new DataView(headerContent.buffer, headerContent.byteOffset, headerContent.byteLength);
        const attr = headerView.getUint32(36, true);
        this.compressed = (attr & 1) !== 0;
    }

    // ─── DocInfo ──────────────────────────────────────────────

    private parseDocInfo() {
        const docInfoEntry = this.ole.FullPaths.find(p => p.toLowerCase().endsWith('docinfo'));
        if (!docInfoEntry) return;

        const docInfoFile = cfb.find(this.ole, docInfoEntry);
        if (!docInfoFile || !docInfoFile.content) return;

        let stream = new Uint8Array(docInfoFile.content as any);
        if (this.compressed) {
            try { stream = pako.inflate(stream, { raw: true }); } catch (_) { }
        }

        const sv = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
        let offset = 0;

        while (offset < stream.length) {
            if (offset + 4 > stream.length) break;
            const header = sv.getUint32(offset, true);
            offset += 4;
            const tagId = header & 0x3FF;
            let size = (header >>> 20) & 0xFFF;

            if (size === 0xFFF) {
                if (offset + 4 > stream.length) break;
                size = sv.getUint32(offset, true);
                offset += 4;
            }

            if (size === 0) { continue; }
            if (offset + size > stream.length) break;

            if (tagId === 17) { // HWPTAG_FACE_NAME (0x11 = 17 = HWPTAG_BEGIN + 1)
                this.parseFaceNameRecord(sv, offset, size);
            } else if (tagId === 21) { // HWPTAG_CHAR_SHAPE (0x15 = 21)
                this.parseCharShapeRecord(sv, offset, size);
            } else if (tagId === 25) { // HWPTAG_PARA_SHAPE (0x19 = 25)
                this.parseParaShapeRecord(sv, offset, size);
            } else if (tagId === 20) { // HWPTAG_BORDER_FILL (0x14 = 20)
                this.parseBorderFillRecord(sv, offset, size, stream.length);
            }

            offset += size;
        }
    }

    // ─── 개별 레코드 파싱 ─────────────────────────────────────

    /** HWPTAG_FACE_NAME(17) - 글꼴 이름 추출 */
    private parseFaceNameRecord(sv: DataView, offset: number, size: number) {
        try {
            if (size < 3) return;
            const nameLen = sv.getUint16(offset + 1, true);
            if (nameLen === 0 || offset + 3 + nameLen * 2 > offset + size) return;
            const nameBytes = new Uint8Array(sv.buffer, sv.byteOffset + offset + 3, nameLen * 2);
            const fontName = new TextDecoder('utf-16le').decode(nameBytes).replace(/\0/g, '');
            this.faceNames.push(fontName);
        } catch (_) { }
    }

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
    private parseCharShapeRecord(sv: DataView, offset: number, size: number) {
        try {
            // 언어별 글꼴 ID
            const fontIds: number[] = [];
            for (let lang = 0; lang < 7; lang++) {
                if (offset + lang * 2 + 2 <= offset + size) {
                    fontIds.push(sv.getUint16(offset + lang * 2, true));
                } else {
                    fontIds.push(0);
                }
            }

            // 자간 (한글 기준: offset+21)
            const letterSpacing = (size >= 22) ? sv.getInt8(offset + 21) : 0;

            // 기준 크기
            let baseSizePt = 10;
            if (size >= 46) {
                const baseSizeInHwpUnit = sv.getInt32(offset + 42, true);
                baseSizePt = Math.round(baseSizeInHwpUnit / 100.0);
            }

            // 속성
            let isBold = false, isItalic = false, isUnderline = false, isStrike = false;
            if (size >= 50) {
                const property = sv.getUint32(offset + 46, true);
                isItalic = (property & 1) !== 0;
                isBold = (property & 2) !== 0;
                isUnderline = (property & 12) !== 0;
                isStrike = (property & (7 << 18)) !== 0;
            }

            // 글자 색
            let charColor = '#000000';
            if (size >= 56) {
                const colorVal = sv.getUint32(offset + 52, true);
                charColor = `#${(colorVal & 0xFF).toString(16).padStart(2, '0')}${((colorVal >> 8) & 0xFF).toString(16).padStart(2, '0')}${((colorVal >> 16) & 0xFF).toString(16).padStart(2, '0')}`;
            }

            this.charShapes.push({ fontIds, isBold, isItalic, isUnderline, isStrike, baseSize: baseSizePt, charColor, letterSpacing });
        } catch (_) { }
    }

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
    private parseParaShapeRecord(sv: DataView, offset: number, size: number) {
        try {
            const prop1 = sv.getUint32(offset, true);
            const lineHeightType = prop1 & 0x3;
            const align = (prop1 >>> 2) & 0x7;

            const leftMargin      = size >= 8  ? sv.getInt32(offset + 4,  true) : 0;
            const rightMargin     = size >= 12 ? sv.getInt32(offset + 8,  true) : 0;
            const indent          = size >= 16 ? sv.getInt32(offset + 12, true) : 0;
            const spaceAbove      = size >= 20 ? sv.getInt32(offset + 16, true) : 0;
            const spaceBelow      = size >= 24 ? sv.getInt32(offset + 20, true) : 0;
            const lineHeightOld   = size >= 28 ? sv.getInt32(offset + 24, true) : 0;

            const borderFillId     = size >= 34 ? sv.getUint16(offset + 32, true) : 0;
            const borderDistLeft   = size >= 36 ? sv.getInt16(offset + 34, true) : 0;
            const borderDistRight  = size >= 38 ? sv.getInt16(offset + 36, true) : 0;
            const borderDistTop    = size >= 40 ? sv.getInt16(offset + 38, true) : 0;
            const borderDistBottom = size >= 42 ? sv.getInt16(offset + 40, true) : 0;

            // 5.0.2.5 이상: 줄 간격 (offset+50, UINT32) 우선 사용
            const lineHeightNew = size >= 54 ? sv.getUint32(offset + 50, true) : 0;
            const lineHeight = lineHeightNew > 0 ? lineHeightNew : lineHeightOld;

            this.paraShapes.push({
                alignment: align, leftMargin, rightMargin, indent,
                lineHeight, lineHeightType,
                spaceAbove, spaceBelow,
                borderFillId,
                borderDistLeft, borderDistRight, borderDistTop, borderDistBottom,
            });
        } catch (_) { }
    }

    /** HWPTAG_BORDER_FILL(20) - 테두리/배경 */
    private parseBorderFillRecord(sv: DataView, offset: number, size: number, streamLen: number) {
        try {
            const parseBorder = (o: number) => {
                if (offset + o + 8 > streamLen) return { type: 0, thickness: 0, color: '#000000' };
                const type = sv.getUint16(offset + o, true);
                const thickness = sv.getUint16(offset + o + 2, true);
                const color = sv.getUint32(offset + o + 4, true);
                return {
                    type, thickness,
                    color: `#${(color & 0xFF).toString(16).padStart(2, '0')}${((color >> 8) & 0xFF).toString(16).padStart(2, '0')}${((color >> 16) & 0xFF).toString(16).padStart(2, '0')}`
                };
            };

            let faceColor: string | undefined;
            const left = parseBorder(2);
            const right = parseBorder(10);
            const top = parseBorder(18);
            const bottom = parseBorder(26);

            if (size >= 50) {
                const fillType = sv.getUint32(offset + 42, true);
                const currentOffset = offset + 46;
                if ((fillType & 0x00000001) !== 0 && currentOffset + 4 <= streamLen) {
                    const colorVal = sv.getUint32(currentOffset, true);
                    if (colorVal !== 0xFFFFFFFF) {
                        const r = colorVal & 0xFF;
                        const g = (colorVal >> 8) & 0xFF;
                        const b = (colorVal >> 16) & 0xFF;
                        faceColor = `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
                    }
                }
            }
            this.borderFills.push({ left, right, top, bottom, faceColor });
        } catch (_) { }
    }

    // ─── 이미지 추출 ──────────────────────────────────────────

    /** OLE BinData 스토리지에서 이미지 파일 추출 */
    private extractImages() {
        const binPaths = this.ole.FullPaths.filter(p => p.toLowerCase().includes('bindata/'));

        for (let i = 0; i < binPaths.length; i++) {
            const binPath = binPaths[i];
            const binEntry = cfb.find(this.ole, binPath);
            if (!binEntry || !binEntry.content) continue;

            let imgData = new Uint8Array(binEntry.content as any);

            // zlib 압축 해제 시도
            try {
                if (imgData[0] === 0x78) {
                    imgData = pako.inflate(imgData);
                }
            } catch (e) {
                console.warn(`이미지 압축 해제 실패 (${binPath}):`, e);
            }

            const match = binPath.match(/\.([a-zA-Z0-9]+)$/);
            const ext = match ? match[1].toLowerCase() : 'png';

            const numMatch = binPath.match(/BIN(\d+)/i);
            const binId = numMatch ? parseInt(numMatch[1], 10) : i + 1;

            this.imageRels.push({
                id: `rIdImg${binId}`,
                target: `media/image${binId}.${ext}`,
                data: imgData,
                ext: ext,
            });
        }
    }
}
