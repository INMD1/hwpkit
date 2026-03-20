import { CharShape, ParaShape, BorderFill, TableCell, PageDef, TextSegment } from './type';

/**
 * DOCX XML 생성 담당
 * HWP에서 파싱된 데이터를 기반으로 DOCX XML 요소를 직접 생성합니다.
 * paragraph, run, table, image 등의 XML을 담당합니다.
 */
export class DocxGenerator {
    private faceNames: string[];
    private charShapes: CharShape[];
    private paraShapes: ParaShape[];
    private pageDefs: PageDef[];
    private borderFills: BorderFill[];
    private nextImageId = 1;

    constructor(
        faceNames: string[],
        charShapes: CharShape[],
        paraShapes: ParaShape[],
        pageDefs: PageDef[],
        borderFills: BorderFill[] = [],
    ) {
        this.faceNames = faceNames;
        this.charShapes = charShapes;
        this.paraShapes = paraShapes;
        this.pageDefs = pageDefs;
        this.borderFills = borderFills;
    }

    // ─── 섹션 파싱 + XML 생성 ──────────────────────────────────

    /** 섹션 하나의 바이너리 스트림을 파싱하여 DOCX body/header/footer XML을 반환 */
    parseSection(stream: Uint8Array): { body: string; header: string | null; footer: string | null } {
        const sv = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
        let offset = 0;
        let bodyXml = '';

        // 문단 상태
        let currentParaShapeId = 0;
        let currentParaTextData: Uint8Array | null = null;
        let currentParaCharShapes: { pos: number; shapeId: number }[] = [];
        let currentParaLineSegs: TextSegment[] = [];

        // 표 상태
        let inTable = false;
        let tableData: (TableCell | null)[][] = [];
        let currentCellParagraphs: string[] = [];
        let currentCellRowIdx = -1;
        let currentCellColIdx = -1;
        let currentCellColSpan = 1;
        let currentCellRowSpan = 1;
        let currentCellBorderFillId = 0;
        let currentCellWidth = 0;
        let currentCellHeight = 0;
        let tableLevel = -1;
        let tableAlignment = 0;
        let tableBorderFillId = 0;

        // 머리말/꼬리말 상태
        let inHeaderFooter: 'header' | 'footer' | null = null;
        let headerFooterLevel = -1;
        let headerBodyXml = '';
        let footerBodyXml = '';

        // 그림 개체 크기
        let currentShapeWidth = 0;
        let currentShapeHeight = 0;

        /** 현재 문단을 DOCX <w:p> XML로 빌드 (표 셀 내부 여부 전달) */
        const buildCurrentParagraphXml = (): string => {
            if (!currentParaTextData) return '';
            const xml = this.buildParagraphXml(currentParaTextData, currentParaCharShapes, currentParaShapeId, inTable, currentParaLineSegs);
            currentParaTextData = null;
            currentParaCharShapes = [];
            currentParaLineSegs = [];
            return xml;
        };

        /** 테이블/이미지/문단 XML을 올바른 버퍼로 라우팅 */
        const emitToBuf = (xml: string) => {
            if (!xml) return;
            if (inTable) { currentCellParagraphs.push(xml); return; }
            if (inHeaderFooter === 'header') { headerBodyXml += xml; return; }
            if (inHeaderFooter === 'footer') { footerBodyXml += xml; return; }
            bodyXml += xml;
        };

        /** 테이블 XML 생성 후 버퍼 라우팅 */
        const flushTable = () => {
            if (tableData.length === 0) return;
            const xml = this.buildTableXml(tableData, tableAlignment, tableBorderFillId);
            tableData = [];
            tableBorderFillId = 0;
            if (inHeaderFooter === 'header') headerBodyXml += xml;
            else if (inHeaderFooter === 'footer') footerBodyXml += xml;
            else bodyXml += xml;
        };

        // ─── 레코드 순회 ──────────────────────────────────────

        while (offset < stream.length) {
            if (offset + 4 > stream.length) break;

            const header = sv.getUint32(offset, true);
            offset += 4;
            const tagId = header & 0x3FF;
            const level = (header >>> 10) & 0x3FF;
            let size = (header >>> 20) & 0xFFF;

            if (size === 0xFFF) {
                if (offset + 4 > stream.length) break;
                size = sv.getUint32(offset, true);
                offset += 4;
            }

            if (size === 0) { continue; }
            if (offset + size > stream.length) break;

            // ── 머리말/꼬리말 종료 감지 ──
            if (inHeaderFooter !== null && level <= headerFooterLevel) {
                emitToBuf(buildCurrentParagraphXml());
                inHeaderFooter = null;
                headerFooterLevel = -1;
            }

            // ── 표 종료 감지 ──
            else if (inTable && level < tableLevel) {
                if (currentCellRowIdx !== -1 && currentCellColIdx !== -1) {
                    const paraXml = buildCurrentParagraphXml();
                    if (paraXml) currentCellParagraphs.push(paraXml);
                    this.fillTableData(tableData, currentCellRowIdx, currentCellColIdx, {
                        text: currentCellParagraphs.join(''),
                        colSpan: currentCellColSpan, rowSpan: currentCellRowSpan,
                        borderFillId: currentCellBorderFillId,
                        width: currentCellWidth, height: currentCellHeight
                    });
                }
                flushTable();
                inTable = false;
                currentCellParagraphs = [];
                currentCellRowIdx = -1;
                currentCellColIdx = -1;
                tableLevel = -1;
            }

            // ── HWPTAG_CTRL_HEADER (71) ── 머리말/꼬리말 컨트롤
            // ctrlId 'head' = bytes [68,65,61,64] → LE uint32 = 0x64616568
            // ctrlId 'foot' = bytes [66,6F,6F,74] → LE uint32 = 0x746F6F66
            else if (tagId === 71) {
                if (size >= 4) {
                    const ctrlId = sv.getUint32(offset, true);
                    if (ctrlId === 0x64616568) {      // 'head' 머리말
                        inHeaderFooter = 'header';
                        headerFooterLevel = level;
                        headerBodyXml = '';
                    } else if (ctrlId === 0x746F6F66) { // 'foot' 꼬리말
                        inHeaderFooter = 'footer';
                        headerFooterLevel = level;
                        footerBodyXml = '';
                    }
                }
            }

            // ── HWPTAG_PAGE_DEF (73) ── 용지 설정
            else if (tagId === 73) {
                if (size >= 40) {
                    this.pageDefs.push({
                        width: Math.round(sv.getUint32(offset, true) / 5),
                        height: Math.round(sv.getUint32(offset + 4, true) / 5),
                        leftMargin: Math.round(sv.getUint32(offset + 8, true) / 5),
                        rightMargin: Math.round(sv.getUint32(offset + 12, true) / 5),
                        topMargin: Math.round(sv.getUint32(offset + 16, true) / 5),
                        bottomMargin: Math.round(sv.getUint32(offset + 20, true) / 5),
                        headerMargin: Math.round(sv.getUint32(offset + 24, true) / 5),
                        footerMargin: Math.round(sv.getUint32(offset + 28, true) / 5),
                    });
                }
            }

            // ── HWPTAG_SHAPE_COMPONENT (76) ── 개체 경계 박스
            // HWP 5.0 사양 구조:
            //   +0  INT32  ctrlId
            //   +4  INT32  xGrpOffset
            //   +8  INT32  yGrpOffset
            //   +12 UINT16 nGrp
            //   +14 INT32  instId
            //   +18 INT32  curWidth  (HWPUNIT)
            //   +22 INT32  curHeight (HWPUNIT)
            else if (tagId === 76) {
                if (size >= 26) {
                    try {
                        const curW = sv.getInt32(offset + 18, true);
                        const curH = sv.getInt32(offset + 22, true);
                        if (curW > 0 && curH > 0) {
                            currentShapeWidth = Math.round(curW / 5);
                            currentShapeHeight = Math.round(curH / 5);
                        }
                    } catch (e) {
                        console.error('HWPTAG_SHAPE_COMPONENT 파싱 에러:', e);
                    }
                }
            }

            // ── HWPTAG_SHAPE_COMPONENT_PICTURE (85) ── 그림 개체
            // 바이트 구조:
            //   BorderLine  12 byte  (offset  0–11)
            //   ImageRect   32 byte  (offset 12–43)  p0~p3 각 INT32×2
            //     p0(top-left):     X@+12, Y@+16
            //     p1(top-right):    X@+20, Y@+24
            //     p2(bottom-right): X@+28, Y@+32
            //     p3(bottom-left):  X@+36, Y@+40
            //   ImageClip   16 byte  (offset 44–59)
            //   Padding      8 byte  (offset 60–67)
            //   PictureInfo  5 byte  (offset 68–72)
            //     +0 Brightness INT8, +1 Contrast INT8, +2 Effect BYTE
            //     +3 BinData ID UINT16  ← offset+71
            else if (tagId === 85) {
                if (size >= 73) {
                    try {
                        const binItemId = sv.getUint16(offset + 71, true);

                        // ImageRect 좌표에서 폭/높이 계산 (HWPUNIT → twips: /5)
                        const rectW = Math.abs(sv.getInt32(offset + 20, true) - sv.getInt32(offset + 12, true));
                        const rectH = Math.abs(sv.getInt32(offset + 40, true) - sv.getInt32(offset + 16, true));

                        // 페이지 컨텐츠 영역 너비 계산 (pageDef 기준, 없으면 A4 기본값)
                        const pd = this.pageDefs.length > 0 ? this.pageDefs[this.pageDefs.length - 1] : null;
                        const maxContentWidth = pd
                            ? Math.max(1, pd.width - pd.leftMargin - pd.rightMargin)
                            : 8306; // A4(11906) - 표준여백 30mm×2(1800×2) twips

                        // 크기 우선순위: ImageRect → SHAPE_COMPONENT → 컨텐츠 폭의 50%(4:3 비율)
                        const fallbackW = Math.round(maxContentWidth / 2);
                        const fallbackH = Math.round(fallbackW * 3 / 4);
                        let widthTwips  = rectW > 0 ? Math.round(rectW / 5) : (currentShapeWidth  > 0 ? currentShapeWidth  : fallbackW);
                        let heightTwips = rectH > 0 ? Math.round(rectH / 5) : (currentShapeHeight > 0 ? currentShapeHeight : fallbackH);

                        // 컨텐츠 폭을 초과하면 종횡비 유지하며 축소
                        if (widthTwips > maxContentWidth) {
                            const ratio = maxContentWidth / widthTwips;
                            widthTwips = maxContentWidth;
                            heightTwips = Math.round(heightTwips * ratio);
                        }

                        emitToBuf(this.buildImageXml(binItemId, widthTwips, heightTwips));
                    } catch (e) {
                        console.error('HWPTAG_SHAPE_COMPONENT_PICTURE 파싱 에러:', e);
                    }
                }
                currentShapeWidth = 0;
                currentShapeHeight = 0;
            }

            // ── HWPTAG_PARA_HEADER (66) ── 문단 헤더
            else if (tagId === 66) {
                emitToBuf(buildCurrentParagraphXml());
                if (size >= 12) {
                    try { currentParaShapeId = sv.getUint32(offset + 8, true); } catch (_) { }
                }
            }

            // ── HWPTAG_PARA_TEXT (67) ── 문단 텍스트
            else if (tagId === 67) {
                currentParaTextData = stream.slice(offset, offset + size);
            }

            // ── HWPTAG_PARA_CHAR_SHAPE (68) ── 문단 글자 모양 매핑
            else if (tagId === 68) {
                currentParaCharShapes = [];
                const count = Math.floor(size / 8);
                for (let i = 0; i < count; i++) {
                    const pos = sv.getUint32(offset + i * 8, true);
                    const shapeId = sv.getUint32(offset + i * 8 + 4, true);
                    currentParaCharShapes.push({ pos, shapeId });
                }
            }
            // ── HWPTAG_PARA_LINE_SEG (69) ── 문단의 줄 레이아웃 정보
            // 각 항목은 36 bytes: INT32×8 + UINT32×1
            else if (tagId === 69) {
                currentParaLineSegs = [];
                const SEG_BYTES = 36;
                const count = Math.floor(size / SEG_BYTES);
                for (let i = 0; i < count; i++) {
                    const base = offset + i * SEG_BYTES;
                    if (base + SEG_BYTES > offset + size) break;
                    currentParaLineSegs.push({
                        textOffset: sv.getInt32(base, true),
                        lineY: sv.getInt32(base + 4, true),
                        lineHeight: sv.getInt32(base + 8, true),
                        textHeight: sv.getInt32(base + 12, true),
                        baselineOffset: sv.getInt32(base + 16, true),
                        lineSpacing: sv.getInt32(base + 20, true),
                        columnStart: sv.getInt32(base + 24, true),
                        segmentWidth: sv.getInt32(base + 28, true),
                        flags: sv.getUint32(base + 32, true),
                    });
                }
            }


            // ── HWPTAG_TABLE (77) ── 표 시작
            else if (tagId === 77) {
                emitToBuf(buildCurrentParagraphXml());

                // 표 속성 파싱
                // +0  UINT32 속성 (bit 0: 글자처럼취급 0=float 1=inline)
                // +20 UINT16 borderFillId
                let tableIsInline = false;
                if (size >= 4) {
                    try {
                        const prop = sv.getUint32(offset, true);
                        tableIsInline = (prop & 1) !== 0;
                    } catch (_) { }
                }
                if (size >= 22) {
                    try {
                        tableBorderFillId = sv.getUint16(offset + 20, true);
                    } catch (_) { }
                }

                // 글자처럼 취급일 때만 현재 문단 정렬을 표 정렬로 적용
                tableAlignment = (tableIsInline && currentParaShapeId < this.paraShapes.length)
                    ? this.paraShapes[currentParaShapeId].alignment
                    : 0;

                inTable = true;
                tableData = [];
                currentCellParagraphs = [];
                currentCellRowIdx = -1;
                currentCellColIdx = -1;
                currentCellBorderFillId = 0;
                currentCellWidth = 0;
                currentCellHeight = 0;
                tableLevel = level;
            }

            // ── HWPTAG_LIST_HEADER (72) ── 표 셀 속성
            else if (tagId === 72) {
                if (inTable) {
                    const paraXml = buildCurrentParagraphXml();
                    if (paraXml) currentCellParagraphs.push(paraXml);

                    if (currentCellRowIdx !== -1 && currentCellColIdx !== -1) {
                        this.fillTableData(tableData, currentCellRowIdx, currentCellColIdx, {
                            text: currentCellParagraphs.join(''),
                            colSpan: currentCellColSpan, rowSpan: currentCellRowSpan,
                            borderFillId: currentCellBorderFillId,
                            width: currentCellWidth, height: currentCellHeight
                        });
                    }
                    currentCellParagraphs = [];

                    if (size >= 26) {
                        try {
                            const attr = offset + 8;
                            currentCellColIdx = sv.getUint16(attr, true);
                            currentCellRowIdx = sv.getUint16(attr + 2, true);
                            currentCellColSpan = sv.getUint16(attr + 4, true);
                            currentCellRowSpan = sv.getUint16(attr + 6, true);
                            currentCellWidth = sv.getUint32(attr + 8, true);
                            currentCellHeight = sv.getUint32(attr + 12, true);
                            if (size >= 34) {
                                currentCellBorderFillId = sv.getUint16(attr + 24, true);
                            }
                        } catch (_) { }
                    }
                }
            }

            offset += size;
        }

        // ── 남은 문단/표 flush ──
        const lastParaXml = buildCurrentParagraphXml();
        if (inTable) {
            if (lastParaXml) currentCellParagraphs.push(lastParaXml);
            if (currentCellRowIdx !== -1 && currentCellColIdx !== -1) {
                this.fillTableData(tableData, currentCellRowIdx, currentCellColIdx, {
                    text: currentCellParagraphs.join(''),
                    colSpan: currentCellColSpan, rowSpan: currentCellRowSpan,
                    borderFillId: currentCellBorderFillId,
                    width: currentCellWidth, height: currentCellHeight
                });
            }
            flushTable();
        } else {
            emitToBuf(lastParaXml);
        }

        return {
            body: bodyXml,
            header: headerBodyXml || null,
            footer: footerBodyXml || null,
        };
    }

    // ─── 문단 XML 빌드 ──────────────────────────────────────────

    /** PARA_TEXT + PARA_CHAR_SHAPE + PARA_SHAPE + LINE_SEG → 완전한 <w:p> */
    buildParagraphXml(
        textData: Uint8Array,
        charShapeMap: { pos: number; shapeId: number }[],
        paraShapeId: number,
        inTableCell = false,
        lineSegs: TextSegment[] = [],
    ): string {
        // LINE_SEG 플래그에서 페이지/단 나누기 위치 추출
        // - 첫 번째 세그먼트에 page break 플래그 → 문단 속성 w:pageBreakBefore
        // - 이후 세그먼트에 page/column break 플래그 → 해당 textOffset에 인라인 <w:br>
        const FLAG_PAGE = 0x10;
        const FLAG_COLUMN = 0x20;

        const pageBreakBefore = lineSegs.length > 0 && (lineSegs[0].flags & FLAG_PAGE) !== 0;

        const inlineBreaks = new Map<number, 'page' | 'column'>();
        for (let si = 1; si < lineSegs.length; si++) {
            const seg = lineSegs[si];
            if (seg.flags & FLAG_PAGE) inlineBreaks.set(seg.textOffset, 'page');
            else if (seg.flags & FLAG_COLUMN) inlineBreaks.set(seg.textOffset, 'column');
        }

        const runs = this.buildRunsFromText(textData, charShapeMap, inlineBreaks);
        if (!runs) return '';

        let xml = '<w:p>';
        xml += this.buildParaPropsXml(paraShapeId, inTableCell, pageBreakBefore);
        xml += runs;
        xml += '</w:p>';
        return xml;
    }

    /** ParaShape → <w:pPr> */
    buildParaPropsXml(paraShapeId: number, inTableCell = false, pageBreakBefore = false): string {
        if (paraShapeId < 0 || paraShapeId >= this.paraShapes.length) return '';

        const ps = this.paraShapes[paraShapeId];
        let pPr = '';

        // 쪽 나누기 후 첫 문단 (LINE_SEG 첫 세그먼트의 page break 플래그)
        if (pageBreakBefore) {
            pPr += '<w:pageBreakBefore/>';
        }

        // 정렬 (0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데, 4=배분, 5=나눔)
        const jc = this.alignToDocx(ps.alignment);
        if (jc !== 'both') {
            pPr += `<w:jc w:val="${jc}"/>`;
        }

        // 들여쓰기 (HWPUNIT → twips: /5)
        // 표 셀 내부에서는 leftMargin/rightMargin을 적용하지 않음 (셀 경계와 중복 방지)
        // 단, hanging이 있으면 w:left = hanging으로 설정하여 첫 줄 위치가 0이 되도록 함
        // (DOCX에서 첫 줄 위치 = w:left - w:hanging)
        {
            const firstLine = ps.indent > 0 ? Math.round(ps.indent / 5) : 0;
            // DOCX w:hanging은 양수여야 함 (내어쓰기 크기의 절대값)
            const hanging = ps.indent < 0 ? Math.abs(Math.round(ps.indent / 5)) : 0;

            let left: number;
            let right: number;
            if (inTableCell) {
                // 표 셀: margin 제외, hanging이 있으면 left = hanging (첫 줄 시작 위치 = 0)
                left = hanging > 0 ? hanging : 0;
                right = 0;
            } else {
                left = Math.max(hanging, Math.round(ps.leftMargin / 5));
                right = Math.max(0, Math.round(ps.rightMargin / 5));
            }
            if (
                left !== undefined ||
                right !== undefined ||
                firstLine !== undefined ||
                hanging !== undefined
            ) {
                let indAttr = '';

                if (left !== undefined) indAttr += ` w:left="${left}"`;
                if (right !== undefined) indAttr += ` w:right="${right}"`;

                if (hanging !== undefined)
                    indAttr += ` w:hanging="${hanging}"`;
                else if (firstLine !== undefined)
                    indAttr += ` w:firstLine="${firstLine}"`;

                pPr += `<w:ind${indAttr}/>`;
            }
        }

        // 문단 간격 + 줄 간격 → 단일 <w:spacing> 요소로 합산
        {
            const before = ps.spaceAbove ? Math.max(0, Math.round(ps.spaceAbove / 5)) : 0;
            const after = ps.spaceBelow ? Math.max(0, Math.round(ps.spaceBelow / 5)) : 0;
            let lineVal = 0;
            let lineRule = '';

            if (ps.lineHeight > 0) {
                if (ps.lineHeightType === 0) {
                    // 글자에 따라 (퍼센트) → auto: 240 = 100%
                    lineVal = Math.round(ps.lineHeight * 240 / 100);
                    lineRule = 'auto';
                } else if (ps.lineHeightType === 1) {
                    // 고정값 (HWPUNIT → twips)
                    lineVal = Math.round((ps.lineHeight / 5));
                    lineRule = 'exact';
                } else if (ps.lineHeightType === 2) {
                    // 여백만 (최소값, HWPUNIT → twips)
                    lineVal = Math.round((ps.lineHeight / 5));
                    lineRule = 'atLeast';
                }
            }
            if (before || after || lineVal) {
                let spacingAttr = '';
                if (before) spacingAttr += ` w:before="${before}"`;
                if (after) spacingAttr += ` w:after="${after}"`;
                if (lineVal) spacingAttr += ` w:line="${lineVal}" w:lineRule="${lineRule}"`;
                pPr += `<w:spacing${spacingAttr}/>`;
            }
        }

        // 문단 테두리/배경 <w:pBdr>, <w:shd>
        if (ps.borderFillId > 0 && ps.borderFillId <= this.borderFills.length) {
            const bf = this.borderFills[ps.borderFillId - 1];

            // 배경색
            if (bf.faceColor) {
                pPr += `<w:shd w:val="clear" w:color="auto" w:fill="${bf.faceColor.replace('#', '')}"/>`;
            }

            // 테두리 (선 종류가 0=없음이 아닌 방향만 출력)
            const sides = [
                { key: 'top',    border: bf.top },
                { key: 'left',   border: bf.left },
                { key: 'bottom', border: bf.bottom },
                { key: 'right',  border: bf.right },
            ] as const;
            const hasAnyBorder = sides.some(s => s.border && s.border.type !== 0);
            if (hasAnyBorder) {
                pPr += '<w:pBdr>';
                for (const { key, border } of sides) {
                    if (border && border.type !== 0) {
                        pPr += this.buildBorderXml(key, border);
                    }
                }
                pPr += '</w:pBdr>';
            }
        }

        if (!pPr) return '';
        return `<w:pPr>${pPr}</w:pPr>`;
    }

    // ─── Run XML 빌드 ──────────────────────────────────────────

    /** PARA_TEXT 바이너리에서 <w:r> 요소들 직접 생성
     * @param inlineBreaks  LINE_SEG에서 추출한 {문자위치 → 'page'|'column'} 맵.
     *                      해당 위치 직전에 <w:br w:type="..."/> 를 삽입한다.
     */
    buildRunsFromText(
        textData: Uint8Array,
        charShapeMap: { pos: number; shapeId: number }[],
        inlineBreaks: Map<number, 'page' | 'column'> = new Map(),
    ): string {
        const dv = new DataView(textData.buffer, textData.byteOffset, textData.byteLength);
        const words = Math.floor(textData.length / 2);
        if (words === 0) return '';

        let result = '';
        let currentText = '';
        let currentShapeId = charShapeMap.length > 0 ? charShapeMap[0].shapeId : -1;
        let shapeIdx = charShapeMap.length > 0 && charShapeMap[0].pos === 0 ? 1 : 0;

        const flushRun = () => {
            if (!currentText) return;
            result += this.buildRunXml(currentText, currentShapeId);
            currentText = '';
        };

        let i = 0;
        while (i < words) {
            // 이 문자 위치 직전에 페이지/단 나누기 삽입 (LINE_SEG 플래그)
            const breakType = inlineBreaks.get(i);
            if (breakType) {
                flushRun();
                result += `<w:r><w:br w:type="${breakType}"/></w:r>`;
            }

            // 글자 모양 전환
            if (shapeIdx < charShapeMap.length && charShapeMap[shapeIdx].pos === i) {
                flushRun();
                currentShapeId = charShapeMap[shapeIdx].shapeId;
                shapeIdx++;
            }

            const chCode = dv.getUint16(i * 2, true);

            if (chCode === 10 || chCode === 13) {
                flushRun();
                i += 1;
            } else if (chCode >= 0 && chCode <= 31) {
                if ([0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31].includes(chCode)) {
                    i += 1;
                } else {
                    i += 8; // 인라인/확장 컨트롤 (표, 그림 등 8-word 항목)
                }
            } else {
                currentText += String.fromCharCode(chCode);
                i += 1;
            }
        }

        flushRun();
        return result;
    }

    /** 텍스트 + CharShape ID → <w:r> XML (서식 포함) */
    buildRunXml(text: string, shapeId: number): string {
        let xml = '<w:r>';

        if (shapeId >= 0 && shapeId < this.charShapes.length) {
            const cs = this.charShapes[shapeId];
            let rPr = '';

            // 글꼴 이름
            if (cs.fontIds && cs.fontIds.length > 0) {
                const koFontId = cs.fontIds[0];
                const enFontId = cs.fontIds[1];
                const koFont = koFontId < this.faceNames.length ? this.faceNames[koFontId] : '';
                const enFont = enFontId < this.faceNames.length ? this.faceNames[enFontId] : '';
                if (koFont || enFont) {
                    const eastAsia = koFont || enFont;
                    const ascii = enFont || koFont;
                    rPr += `<w:rFonts w:ascii="${this.escapeXml(ascii)}" w:eastAsia="${this.escapeXml(eastAsia)}" w:hAnsi="${this.escapeXml(ascii)}"/>`;
                }
            }

            if (cs.isBold) rPr += '<w:b/><w:bCs/>';
            if (cs.isItalic) rPr += '<w:i/><w:iCs/>';
            if (cs.isUnderline) rPr += '<w:u w:val="single"/>';
            if (cs.isStrike) rPr += '<w:strike/>';

            if (cs.baseSize > 0) {
                const sz = cs.baseSize * 2;
                rPr += `<w:sz w:val="${sz}"/>`;
                rPr += `<w:szCs w:val="${sz}"/>`;
            }

            if (cs.charColor && cs.charColor !== '#000000') {
                rPr += `<w:color w:val="${cs.charColor.replace('#', '')}"/>`;
            }

            if (cs.letterSpacing && cs.letterSpacing !== 0) {
                const spacingVal = Math.round(cs.letterSpacing);
                rPr += `<w:spacing w:val="${spacingVal}"/>`;
            }

            if (rPr) {
                xml += `<w:rPr>${rPr}</w:rPr>`;
            }
        }

        xml += `<w:t xml:space="preserve">${this.escapeXml(text)}</w:t>`;
        xml += '</w:r>';
        return xml;
    }

    // ─── 테이블 XML 빌드 ────────────────────────────────────────

    /** tableData 2D 배열 → 완전한 <w:tbl> XML */
    buildTableXml(tableData: (TableCell | null)[][], alignment = 0, tableBorderFillId = 0): string {
        let maxCols = 0;
        for (const row of tableData) {
            if (row.length > maxCols) maxCols = row.length;
        }
        if (maxCols === 0) return '';

        // ── 1. 다중 패스로 컬럼 폭 계산 ───────────────────────────────
        const colWidths: number[] = new Array(maxCols).fill(0);

        // Pass 1: colSpan=1인 셀에서 정확한 컬럼 폭 수집
        for (const row of tableData) {
            let c = 0;
            while (c < maxCols) {
                const cell = c < row.length ? row[c] : null;
                if (cell && !cell.isVMergeContinue) {
                    const span = Math.max(1, cell.colSpan || 1);
                    if (span === 1 && cell.width > 0) {
                        const w = Math.round(cell.width / 5);
                        if (w > colWidths[c]) colWidths[c] = w;
                    }
                    c += span;
                } else {
                    c++;
                }
            }
        }

        // Pass 2: colSpan>1인 셀에서 아직 0인 컬럼에 폭 분배
        for (const row of tableData) {
            let c = 0;
            while (c < maxCols) {
                const cell = c < row.length ? row[c] : null;
                if (cell && !cell.isVMergeContinue) {
                    const span = Math.min(Math.max(1, cell.colSpan || 1), maxCols - c);
                    if (span > 1 && cell.width > 0) {
                        const totalTwips = Math.round(cell.width / 5);
                        let zeroCount = 0, knownSum = 0;
                        for (let s = 0; s < span; s++) {
                            if (colWidths[c + s] === 0) zeroCount++;
                            else knownSum += colWidths[c + s];
                        }
                        if (zeroCount > 0) {
                            const perUnset = Math.max(1, Math.round((totalTwips - knownSum) / zeroCount));
                            for (let s = 0; s < span; s++) {
                                if (colWidths[c + s] === 0) colWidths[c + s] = perUnset;
                            }
                        }
                    }
                    c += span;
                } else {
                    c++;
                }
            }
        }

        // Pass 3: 여전히 0인 컬럼은 기본값으로 채우기
        for (let c = 0; c < maxCols; c++) {
            if (colWidths[c] === 0) colWidths[c] = Math.round(9000 / maxCols);
        }

        const totalWidth = colWidths.reduce((s, w) => s + w, 0);

        // ── 2. 행 높이 계산 ────────────────────────────────────────────
        const rowHeights: number[] = new Array(tableData.length).fill(0);
        for (let r = 0; r < tableData.length; r++) {
            const row = tableData[r];
            // rowSpan=1 셀에서 최대 높이 수집
            for (const cell of row) {
                if (cell && !cell.isVMergeContinue && (cell.rowSpan || 1) === 1 && cell.height > 0) {
                    const h = Math.round(cell.height / 5);
                    if (h > rowHeights[r]) rowHeights[r] = h;
                }
            }
            // rowSpan=1 셀이 없으면 rowSpan>1 셀에서 1행 분할
            if (rowHeights[r] === 0) {
                for (const cell of row) {
                    if (cell && !cell.isVMergeContinue && cell.height > 0) {
                        rowHeights[r] = Math.round(cell.height / Math.max(1, cell.rowSpan || 1) / 5);
                        break;
                    }
                }
            }
        }

        // ── 3. XML 빌드 ────────────────────────────────────────────────
        let tblXml = '<w:tbl>';
        tblXml += '<w:tblPr>';
        const tblJc = this.alignToDocx(alignment);
        if (tblJc !== 'both') tblXml += `<w:jc w:val="${tblJc}"/>`;
        tblXml += `<w:tblW w:w="${totalWidth}" w:type="dxa"/><w:tblBorders>`;

        // 표 전체 테두리: tableBorderFillId가 있으면 해당 스타일 사용
        if (tableBorderFillId > 0 && tableBorderFillId <= this.borderFills.length) {
            const bf = this.borderFills[tableBorderFillId - 1];
            tblXml += this.buildBorderXml('top', bf.top);
            tblXml += this.buildBorderXml('left', bf.left);
            tblXml += this.buildBorderXml('bottom', bf.bottom);
            tblXml += this.buildBorderXml('right', bf.right);
            tblXml += this.buildBorderXml('insideH', bf.top);
            tblXml += this.buildBorderXml('insideV', bf.left);
        } else {
            tblXml += '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
            tblXml += '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
            tblXml += '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
            tblXml += '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
            tblXml += '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
            tblXml += '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        }

        tblXml += '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>';

        // tblGrid
        tblXml += '<w:tblGrid>';
        for (let c = 0; c < maxCols; c++) {
            tblXml += `<w:gridCol w:w="${colWidths[c]}"/>`;
        }
        tblXml += '</w:tblGrid>';

        // 행/셀
        for (let r = 0; r < tableData.length; r++) {
            tblXml += '<w:tr>';
            if (rowHeights[r] > 0) {
                tblXml += `<w:trPr><w:trHeight w:val="${rowHeights[r]}" w:hRule="atLeast"/></w:trPr>`;
            }
            for (let c = 0; c < maxCols; c++) {
                const cellObj = tableData[r][c];
                if (!cellObj) continue;

                tblXml += '<w:tc><w:tcPr>';

                // 셀 폭: 스패닝된 컬럼 폭의 합 (tblGrid와 일관성 유지)
                const span = Math.min(Math.max(1, cellObj.colSpan || 1), maxCols - c);
                let cellWidthTwips = 0;
                for (let s = 0; s < span; s++) {
                    cellWidthTwips += colWidths[c + s] || 0;
                }
                tblXml += `<w:tcW w:w="${cellWidthTwips}" w:type="dxa"/>`;

                if ((cellObj.colSpan || 1) > 1) tblXml += `<w:gridSpan w:val="${cellObj.colSpan}"/>`;

                if ((cellObj.rowSpan || 1) > 1 && !cellObj.isVMergeContinue) {
                    tblXml += '<w:vMerge w:val="restart"/>';
                } else if (cellObj.isVMergeContinue) {
                    tblXml += '<w:vMerge/>';
                }

                // 셀 테두리 스타일
                if (cellObj.borderFillId > 0 && cellObj.borderFillId <= this.borderFills.length) {
                    const bf = this.borderFills[cellObj.borderFillId - 1];
                    tblXml += '<w:tcBorders>';
                    tblXml += this.buildBorderXml('top', bf.top);
                    tblXml += this.buildBorderXml('left', bf.left);
                    tblXml += this.buildBorderXml('bottom', bf.bottom);
                    tblXml += this.buildBorderXml('right', bf.right);
                    tblXml += '</w:tcBorders>';
                    // 배경색
                    if (bf.faceColor) {
                        tblXml += `<w:shd w:val="clear" w:color="auto" w:fill="${bf.faceColor.replace('#', '')}"/>`;
                    }
                }

                tblXml += '</w:tcPr>';

                if (cellObj.isVMergeContinue) {
                    tblXml += '<w:p/>';
                } else {
                    const cellContent = cellObj.text;
                    tblXml += cellContent ? cellContent : '<w:p/>';
                }

                tblXml += '</w:tc>';

                if ((cellObj.colSpan || 1) > 1) {
                    c += (cellObj.colSpan - 1);
                }
            }
            tblXml += '</w:tr>';
        }

        tblXml += '</w:tbl>';
        return tblXml;
    }

    /** 2D 배열에 셀 데이터 채우기 (병합 처리 포함) */
    fillTableData(tableData: (TableCell | null)[][], rIdx: number, cIdx: number, cell: TableCell) {
        while (tableData.length <= rIdx + cell.rowSpan - 1) tableData.push([]);

        while (tableData[rIdx].length <= cIdx) tableData[rIdx].push(null);
        tableData[rIdx][cIdx] = cell;

        if (cell.rowSpan > 1) {
            for (let i = 1; i < cell.rowSpan; i++) {
                while (tableData[rIdx + i].length <= cIdx) tableData[rIdx + i].push(null);
                tableData[rIdx + i][cIdx] = { ...cell, text: '', isVMergeContinue: true };
            }
        }

        if (cell.colSpan > 1) {
            for (let r = 0; r < cell.rowSpan; r++) {
                for (let c = 1; c < cell.colSpan; c++) {
                    while (tableData[rIdx + r].length <= cIdx + c) tableData[rIdx + r].push(null);
                }
            }
        }
    }

    // ─── 이미지 XML ──────────────────────────────────────────

    /** DOCX DrawingML 이미지 요소 생성 */
    buildImageXml(binDataId: number, widthTwips: number, heightTwips: number): string {
        const imgId = this.nextImageId++;
        const relId = `rIdImg${binDataId}`;
        const cxEmu = Math.round(widthTwips * 635);
        const cyEmu = Math.round(heightTwips * 635);
        return `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cxEmu}" cy="${cyEmu}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${imgId}" name="Picture ${imgId}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${imgId}" name="Image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cxEmu}" cy="${cyEmu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    }

    // ─── 구역/페이지 설정 ────────────────────────────────────────

    /** 구역 설정 <w:sectPr> XML */
    buildSectPrXml(hasHeader = false, hasFooter = false): string {
        const pd = this.pageDefs.length > 0 ? this.pageDefs[this.pageDefs.length - 1] : null;
        let s = '<w:sectPr>';
        if (hasHeader) s += '<w:headerReference w:type="default" r:id="rIdHdr1"/>';
        if (hasFooter) s += '<w:footerReference w:type="default" r:id="rIdFtr1"/>';
        if (pd) {
            s += `<w:pgSz w:w="${pd.width}" w:h="${pd.height}"/>`;
            s += `<w:pgMar w:top="${pd.topMargin}" w:right="${pd.rightMargin}" w:bottom="${pd.bottomMargin}" w:left="${pd.leftMargin}" w:header="${pd.headerMargin}" w:footer="${pd.footerMargin}" w:gutter="0"/>`;
        } else {
            s += '<w:pgSz w:w="11906" w:h="16838"/>';
            s += '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="851" w:gutter="0"/>';
        }
        s += '</w:sectPr>';
        return s;
    }

    // ─── XML 유틸리티 ────────────────────────────────────────────

    escapeXml(text: string): string {
        return text
            .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    private alignToDocx(align: number): string {
        switch (align) {
            case 1: return 'left';
            case 2: return 'right';
            case 3: return 'center';
            case 4: return 'distribute';
            case 5: return 'thaiDistribute';
            default: return 'both';
        }
    }

    private hwpBorderTypeToDocx(type: number): string {
        switch (type) {
            case 0: return 'nil';
            case 1: return 'single';
            case 2: return 'dashed';
            case 3: return 'dotted';
            case 4: return 'dotDash';
            case 5: return 'dotDotDash';
            case 6: return 'double';
            case 7: return 'thick';
            default: return 'single';
        }
    }

    /** HWP 선 굵기 코드 → DOCX w:sz (1/8pt) */
    private hwpThicknessToDocxSz(code: number): number {
        // HWP 두께 코드: 0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0 (mm)
        // DOCX w:sz = mm * 2.8346pt/mm * 8(1/8pt) ≈ mm * 22.68
        const sizes = [2, 3, 3, 5, 6, 7, 9, 11, 14, 16, 23, 34, 45, 68, 91, 113];
        return sizes[code] ?? 4;
    }

    /** BorderFill의 한 방향 → DOCX <w:${side}> XML 문자열 */
    private buildBorderXml(side: string, border: { type: number; thickness: number; color: string }): string {
        if (!border || border.type === 0) {
            return `<w:${side} w:val="nil"/>`;
        }
        const val = this.hwpBorderTypeToDocx(border.type);
        const sz = this.hwpThicknessToDocxSz(border.thickness);
        const color = border.color.replace('#', '');
        return `<w:${side} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${color}"/>`;
    }
}
