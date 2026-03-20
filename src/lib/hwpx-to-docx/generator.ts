import { ConversionContext } from './context';

export class DocxGenerator {
    private ctx: ConversionContext;

    constructor(ctx: ConversionContext) {
        this.ctx = ctx;
    }

    public convertSection(xml: string): string {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'text/xml');
        const root = doc.documentElement;
        let bodyContent = '';

        this.extractPageSetup(root);

        const children = root.childNodes;
        for (let i = 0; i < children.length; i++) {
            const child = children[i];
            if (child.nodeType !== 1) continue;
            if ((child as Element).localName === 'p') {
                bodyContent += this.convertParagraph(child as Element);
            } else if ((child as Element).localName === 'tbl') {
                bodyContent += this.convertTable(child as Element);
            }
        }

        return bodyContent;
    }

    private extractPageSetup(secRoot: Element) {
        // <hp:secPr> → <hp:pagePr> → 속성(width, height) + 자식 <hp:margin>
        // secPr는 section 루트 또는 그 하위 어딘가에 있을 수 있음
        const pagePrList = secRoot.getElementsByTagNameNS(this.ctx.ns.hp, 'pagePr');
        if (!pagePrList || pagePrList.length === 0) return;
        const pagePr = pagePrList[0];

        // 페이지 크기 (HWPUnit → dxa: ÷5)
        const wHwp = parseInt(pagePr.getAttribute('width') || '0', 10);
        const hHwp = parseInt(pagePr.getAttribute('height') || '0', 10);
        if (wHwp > 0) this.ctx.pageWidth = Math.round(wHwp / 5);
        if (hHwp > 0) this.ctx.pageHeight = Math.round(hHwp / 5);

        // <hp:margin> 자식에서 여백 읽기
        const marginElem = Array.from(pagePr.childNodes)
            .find(n => n.nodeType === 1 && (n as Element).localName === 'margin') as Element | undefined;
        if (!marginElem) return;

        const toInt = (attr: string) => parseInt(marginElem.getAttribute(attr) || '0', 10);
        const topHwp    = toInt('top');
        const bottomHwp = toInt('bottom');
        const leftHwp   = toInt('left');
        const rightHwp  = toInt('right');
        const headerHwp = toInt('header');
        const footerHwp = toInt('footer');

        if (topHwp > 0)    this.ctx.marginTop    = Math.round(topHwp / 5);
        if (bottomHwp > 0) this.ctx.marginBottom = Math.round(bottomHwp / 5);
        if (leftHwp > 0)   this.ctx.marginLeft   = Math.round(leftHwp / 5);
        if (rightHwp > 0)  this.ctx.marginRight  = Math.round(rightHwp / 5);

        // header/footer 여백: HWPX에서 0이면 top/bottom과 동일하게 설정
        this.ctx.marginHeader = headerHwp > 0 ? Math.round(headerHwp / 5) : this.ctx.marginTop;
        this.ctx.marginFooter = footerHwp > 0 ? Math.round(footerHwp / 5) : this.ctx.marginBottom;
    }

    private convertParagraph(pElem: Element): string {
        const paraPrIDRef = pElem.getAttribute('paraPrIDRef') || '0';
        const pageBreak = pElem.getAttribute('pageBreak') || '0';

        let xml = '<w:p>';

        // w:pPr 생성
        xml += this.getParaPropertiesXml(paraPrIDRef, pageBreak);

        // hp:run 순회
        const runs = pElem.getElementsByTagNameNS(this.ctx.ns.hp, 'run');
        for (const run of Array.from(runs)) {
            // run이 현재 p의 직접 자식인지 확인 (subList 내부 run 제외)
            if (run.parentNode !== pElem) continue;

            const charPrIDRef = run.getAttribute('charPrIDRef') || '0';

            // 표 처리 (방어: 현재 run에 직접 속한 최상위 표만 처리, 중첩 표 중복 방지)
            const allTables = Array.from(run.getElementsByTagNameNS(this.ctx.ns.hp, 'tbl'));
            const tables = allTables.filter(tbl => {
                let p = tbl.parentNode;
                while (p && p !== run) {
                    if ((p as Element).localName === 'run') return false;
                    p = p.parentNode;
                }
                return p === run;
            });

            if (tables.length > 0) {
                xml += '</w:p>';
                for (const tbl of tables) {
                    xml += this.convertTable(tbl);
                }
                xml += '<w:p>';
                xml += this.getParaPropertiesXml(paraPrIDRef, '0');
                continue;
            }

            // 이미지 처리
            const pics = run.getElementsByTagNameNS(this.ctx.ns.hp, 'pic');
            if (pics.length > 0) {
                for (const pic of Array.from(pics)) {
                    xml += this.convertImage(pic, charPrIDRef);
                }
                continue;
            }

            // 도형(rect) 내 텍스트 추출
            const rects = run.getElementsByTagNameNS(this.ctx.ns.hp, 'rect');
            if (rects.length > 0) {
                for (const rect of Array.from(rects)) {
                    xml += this.convertRect(rect, charPrIDRef);
                }
                continue;
            }

            // secPr, ctrl은 변환 대상이 아님 (skip)
            const secPr = run.getElementsByTagNameNS(this.ctx.ns.hp, 'secPr');
            if (secPr.length > 0) continue;

            // 일반 텍스트 처리
            const tElems = run.getElementsByTagNameNS(this.ctx.ns.hp, 't');
            if (tElems.length > 0) {
                for (const tElem of Array.from(tElems)) {
                    const text = this.getTextContent(tElem);
                    if (text) {
                        xml += '<w:r>';
                        xml += this.getRunPropertiesXml(charPrIDRef);
                        xml += `<w:t xml:space="preserve">${this.escapeXml(text)}</w:t>`;
                        xml += '</w:r>';
                    }

                    // 탭 처리
                    const tabs = tElem.getElementsByTagNameNS(this.ctx.ns.hp, 'tab');
                    for (const _tab of Array.from(tabs)) {
                        xml += '<w:r>';
                        xml += this.getRunPropertiesXml(charPrIDRef);
                        xml += '<w:tab/>';
                        xml += '</w:r>';
                    }
                }
            // 텍스트/이미지/도형이 없는 빈 run은 생략 (w:r은 반드시 w:t를 포함해야 함)
            }
        }

        xml += '</w:p>';
        return xml;
    }

    private getTextContent(tElem: Element): string {
        let text = '';
        for (const node of Array.from(tElem.childNodes)) {
            if (node.nodeType === 3) { // 텍스트 노드
                text += node.nodeValue;
            }
        }
        // trim() 제거: xml:space="preserve"로 출력하므로 앞뒤 공백을 보존해야 함
        return text;
    }

    private getParaPropertiesXml(paraPrId: string, pageBreak: string): string {
        const pp = this.ctx.paraProperties.get(paraPrId);
        if (!pp && pageBreak !== '1') return '';

        let xml = '<w:pPr>';

        if (pp) {
            // 정렬
            const alignMap: { [key: string]: string } = { 'LEFT': 'left', 'CENTER': 'center', 'RIGHT': 'right', 'JUSTIFY': 'both', 'DISTRIBUTE': 'distribute' };
            if (pp.align && pp.align !== 'LEFT') {
                xml += `<w:jc w:val="${alignMap[pp.align] || 'left'}"/>`;
            }

            // 들여쓰기 (HWPUNIT → twips: ÷5)
            if (pp.leftMargin || pp.rightMargin || pp.indent) {
                let indXml = '<w:ind';
                if (pp.leftMargin) indXml += ` w:left="${Math.round(pp.leftMargin / 5)}"`;
                if (pp.rightMargin) indXml += ` w:right="${Math.round(pp.rightMargin / 5)}"`;
                if (pp.indent > 0) indXml += ` w:firstLine="${Math.round(pp.indent / 5)}"`;
                else if (pp.indent < 0) indXml += ` w:hanging="${Math.round(Math.abs(pp.indent) / 5)}"`;
                indXml += '/>';
                xml += indXml;
            }

            // 줄간격
            let spacingXml = '<w:spacing';
            if (pp.prevSpacing) spacingXml += ` w:before="${Math.round(pp.prevSpacing / 5)}"`;
            if (pp.nextSpacing) spacingXml += ` w:after="${Math.round(pp.nextSpacing / 5)}"`;
            else spacingXml += ' w:after="0"';

            if (pp.lineSpacingType === 'PERCENT') {
                spacingXml += ` w:line="${Math.round(pp.lineSpacingVal * 240 / 100)}"`;
                spacingXml += ' w:lineRule="auto"';
            } else if (pp.lineSpacingType === 'FIXED') {
                spacingXml += ` w:line="${Math.round(pp.lineSpacingVal / 5)}"`;
                spacingXml += ' w:lineRule="exact"';
            } else if (pp.lineSpacingType === 'BETWEENLINES' || pp.lineSpacingType === 'AT_LEAST') {
                spacingXml += ` w:line="${Math.round(pp.lineSpacingVal / 5)}"`;
                spacingXml += ' w:lineRule="atLeast"';
            }
            spacingXml += '/>';
            xml += spacingXml;

            // keepWithNext, keepLines
            if (pp.keepWithNext === '1') xml += '<w:keepNext/>';
            if (pp.keepLines === '1') xml += '<w:keepLines/>';
        }

        // 페이지 나누기
        if (pageBreak === '1') {
            xml += '<w:pageBreakBefore/>';
        }

        xml += '</w:pPr>';
        return xml;
    }

    private getRunPropertiesXml(charPrId: string): string {
        const cp = this.ctx.charProperties.get(charPrId);
        if (!cp) return '';

        let xml = '<w:rPr>';

        // 폰트
        if (cp.latinFont || cp.hangulFont) {
            xml += `<w:rFonts w:ascii="${cp.latinFont}" w:eastAsia="${cp.hangulFont}" w:hAnsi="${cp.latinFont}"/>`;
        }

        // 굵게
        if (cp.bold) xml += '<w:b/>';

        // 기울임
        if (cp.italic) xml += '<w:i/>';

        // 밑줄
        if (cp.underline) xml += '<w:u w:val="single"/>';

        // 취소선
        if (cp.strikeout) xml += '<w:strike/>';

        // 위/아래 첨자
        if (cp.supscript) xml += '<w:vertAlign w:val="superscript"/>';
        if (cp.subscript) xml += '<w:vertAlign w:val="subscript"/>';

        // 글자 크기 (HWPX height → half-points: ÷50)
        const szVal = Math.round(cp.height / 50);
        if (szVal > 0) {
            xml += `<w:sz w:val="${szVal}"/>`;
            xml += `<w:szCs w:val="${szVal}"/>`;
        }

        // 글자 색상
        if (cp.textColor && cp.textColor !== '#000000') {
            xml += `<w:color w:val="${cp.textColor.replace('#', '')}"/>`;
        }

        // 형광색 (shadeColor 또는 borderFill faceColor)
        if (cp.shadeColor && cp.shadeColor !== 'none') {
            xml += `<w:highlight w:val="${this.colorToHighlightName(cp.shadeColor)}"/>`;
        } else if (cp.borderFillIDRef && cp.borderFillIDRef !== '1') {
            const bf = this.ctx.borderFills.get(cp.borderFillIDRef);
            if (bf && bf.faceColor && bf.faceColor !== '#FFFFFF' && bf.faceColor !== 'none') {
                xml += `<w:shd w:val="clear" w:color="auto" w:fill="${bf.faceColor.replace('#', '')}"/>`;
            }
        }

        xml += '</w:rPr>';
        return xml;
    }

    private colorToHighlightName(color: string): string {
        const map: { [key: string]: string } = {
            '#FFFF00': 'yellow', '#00FFFF': 'cyan', '#00FF00': 'green',
            '#FF00FF': 'magenta', '#0000FF': 'blue', '#FF0000': 'red',
            '#000000': 'black', '#FFFFFF': 'white', '#008000': 'darkGreen',
            '#808000': 'darkYellow', '#800000': 'darkRed', '#008080': 'darkCyan',
            '#800080': 'darkMagenta', '#00008B': 'darkBlue', '#A9A9A9': 'darkGray',
            '#D3D3D3': 'lightGray'
        };
        return map[color.toUpperCase()] || 'yellow';
    }

    private convertTable(tblElem: Element): string {
        let xml = '<w:tbl>';

        // ─────────────────────────────────────────────────────────────────
        // 단위 변환 규칙:
        //   HWPUnit (1/7200 inch)  →  dxa (1/1440 inch)
        //   dxa = HWPUnit / 5
        // ─────────────────────────────────────────────────────────────────

        // 1. 표 전체 너비 파싱 (<hp:sz width="..." .../>)
        const szElem = tblElem.getElementsByTagNameNS(this.ctx.ns.hp, 'sz')[0];
        const tblWidthHwp = szElem ? parseInt(szElem.getAttribute('width') || '0', 10) : 0;
        const tblWidthDxa = tblWidthHwp > 0 ? Math.round(tblWidthHwp / 5) : 0;

        // 2. <hp:pos> 정렬 매핑
        //    - treatAsChar="1" : 글자처럼 취급(인라인) → 왼쪽 정렬 유지
        //    - treatAsChar="0" : 플로팅 → horzAlign 속성("LEFT"|"CENTER"|"RIGHT")으로 결정
        //    ※ 과거 코드의 'posRelToPage' 는 존재하지 않는 속성명이었음 → 수정
        let jcVal = 'left';
        const posElem = tblElem.getElementsByTagNameNS(this.ctx.ns.hp, 'pos')[0];
        if (posElem && posElem.getAttribute('treatAsChar') !== '1') {
            const horzAlign = posElem.getAttribute('horzAlign') || 'LEFT';
            if (horzAlign === 'CENTER') jcVal = 'center';
            else if (horzAlign === 'RIGHT') jcVal = 'right';
            else jcVal = 'left';
        }

        // 3. <w:tblPr>
        //    표 외곽 테두리: tbl 의 borderFillIDRef → borderFill 맵 조회
        const tblBfId = tblElem.getAttribute('borderFillIDRef') || '';
        const tblBf = this.ctx.borderFills.get(tblBfId);

        xml += '<w:tblPr>';
        xml += tblWidthDxa > 0
            ? `<w:tblW w:w="${tblWidthDxa}" w:type="dxa"/>`
            : '<w:tblW w:w="0" w:type="auto"/>';
        xml += `<w:jc w:val="${jcVal}"/>`;
        xml += this.buildTblBordersXml(tblBf);
        xml += '<w:tblLayout w:type="fixed"/>';
        xml += '</w:tblPr>';

        // 4. 직접 자식 <hp:tr> 수집
        //    ※ getElementsByTagNameNS 는 중첩 표의 tr까지 반환하므로 childNodes 필터로 대체
        const allRows = Array.from(tblElem.childNodes)
            .filter(n => n.nodeType === 1 && (n as Element).localName === 'tr') as Element[];

        // 5. <w:tblGrid> 구성
        //    모든 행의 셀을 순회하여 <hp:cellAddr colAddr="...">와
        //    tc 직접 자식 <hp:cellSz> / <hp:sz> 를 기준으로 열별 너비 맵 구성
        //
        //    - 셀 너비 소스 : tc 직접 자식 <hp:cellSz> 우선, 없으면 <hp:sz>
        //      (getElementsByTagNameNS 는 중첩 표 요소까지 포함하므로 사용 금지)
        //    - 1단계: colSpan=1 셀로 열 너비를 우선 확정
        //    - 2단계: 아직 너비 미확정 열은 병합 셀의 남은 너비로 균등 분배
        //    - 방어 로직: 합이 0이면 tblWidthDxa 기준 균등 분배
        interface CellInfo { colAddr: number; colSpan: number; widthDxa: number; }
        const allCellInfos: CellInfo[] = [];

        for (const trElem of allRows) {
            const cells = Array.from(trElem.childNodes)
                .filter(n => n.nodeType === 1 && (n as Element).localName === 'tc') as Element[];

            for (const tc of cells) {
                // tc 직접 자식만 검색 (중첩 표 혼입 방지)
                const addrElem = Array.from(tc.childNodes)
                    .find(n => n.nodeType === 1 && (n as Element).localName === 'cellAddr') as Element | undefined;
                const spanElem = Array.from(tc.childNodes)
                    .find(n => n.nodeType === 1 && (n as Element).localName === 'cellSpan') as Element | undefined;
                const cellSzElem = (
                    Array.from(tc.childNodes).find(n => n.nodeType === 1 && (n as Element).localName === 'cellSz') ||
                    Array.from(tc.childNodes).find(n => n.nodeType === 1 && (n as Element).localName === 'sz')
                ) as Element | undefined;

                const colAddr = addrElem ? parseInt(addrElem.getAttribute('colAddr') || '0', 10) : 0;
                const colSpan = spanElem ? Math.max(1, parseInt(spanElem.getAttribute('colSpan') || '1', 10)) : 1;
                const widthHwp = cellSzElem ? parseInt(cellSzElem.getAttribute('width') || '0', 10) : 0;
                const widthDxa = Math.round(widthHwp / 5);

                allCellInfos.push({ colAddr, colSpan, widthDxa });
            }
        }

        // 총 열 수 계산
        let totalColCount = 0;
        for (const ci of allCellInfos) {
            totalColCount = Math.max(totalColCount, ci.colAddr + ci.colSpan);
        }

        // 열 너비 맵 (colAddr → dxa)
        const colWidthMap = new Map<number, number>();

        // 1단계: colSpan=1 셀 우선 확정
        for (const ci of allCellInfos) {
            if (ci.colSpan === 1 && ci.widthDxa > 0 && !colWidthMap.has(ci.colAddr)) {
                colWidthMap.set(ci.colAddr, ci.widthDxa);
            }
        }

        // 2단계: 병합 셀로 미확정 열 보완
        for (const ci of allCellInfos) {
            if (ci.colSpan > 1 && ci.widthDxa > 0) {
                const unknownCols: number[] = [];
                let knownSum = 0;
                for (let c = ci.colAddr; c < ci.colAddr + ci.colSpan; c++) {
                    if (colWidthMap.has(c)) {
                        knownSum += colWidthMap.get(c)!;
                    } else {
                        unknownCols.push(c);
                    }
                }
                if (unknownCols.length > 0) {
                    const remaining = Math.max(0, ci.widthDxa - knownSum);
                    const perUnknown = Math.round(remaining / unknownCols.length);
                    for (const c of unknownCols) {
                        colWidthMap.set(c, perUnknown);
                    }
                }
            }
        }

        // tblGrid 배열 구성
        const tblGridWidths: number[] = [];
        if (totalColCount > 0) {
            for (let c = 0; c < totalColCount; c++) {
                tblGridWidths.push(colWidthMap.get(c) || 0);
            }
        }

        // 방어 로직: 합이 0이면 tblWidthDxa 기준 균등 분배
        const gridTotal = tblGridWidths.reduce((a, b) => a + b, 0);
        if (gridTotal === 0 && tblWidthDxa > 0) {
            const cols = tblGridWidths.length || 1;
            const perCol = Math.round(tblWidthDxa / cols);
            if (tblGridWidths.length === 0) tblGridWidths.push(tblWidthDxa);
            else tblGridWidths.fill(perCol);
        } else if (tblGridWidths.length === 0 && tblWidthDxa > 0) {
            tblGridWidths.push(tblWidthDxa);
        }

        xml += '<w:tblGrid>';
        for (const w of tblGridWidths) xml += `<w:gridCol w:w="${w}"/>`;
        xml += '</w:tblGrid>';

        // 6. 행/셀 변환
        //    rowSpan 처리: HWPX에서는 병합 시작 셀에만 rowSpan이 있고,
        //    계속(continuation) 셀은 XML에 없는 경우가 많음.
        //    → 점유 맵을 활용해 DOCX가 요구하는 vMerge continue 셀을 삽입
        const rowspanOccupied = new Set<string>(); // "r,c" 좌표
        const totalCols = tblGridWidths.length;

        for (let r = 0; r < allRows.length; r++) {
            const trElem = allRows[r];
            const cells = Array.from(trElem.childNodes)
                .filter(n => n.nodeType === 1 && (n as Element).localName === 'tc') as Element[];

            xml += '<w:tr>';

            // 행 높이: 첫 번째 셀의 cellSz.height (HWPUnit → dxa: ÷5)
            if (cells.length > 0) {
                const firstCellSz = Array.from(cells[0].childNodes)
                    .find(n => n.nodeType === 1 && (n as Element).localName === 'cellSz') as Element | undefined;
                const heightHwp = firstCellSz ? parseInt(firstCellSz.getAttribute('height') || '0', 10) : 0;
                if (heightHwp > 0) {
                    xml += `<w:trPr><w:trHeight w:val="${Math.round(heightHwp / 5)}"/></w:trPr>`;
                }
            }

            let colCursor = 0;
            let cellIdx = 0;

            while (colCursor < totalCols || cellIdx < cells.length) {
                // 그리드 범위를 넘는 초과 셀은 무시
                if (colCursor >= totalCols) { cellIdx++; continue; }

                if (rowspanOccupied.has(`${r},${colCursor}`)) {
                    // rowSpan으로 점유된 위치 → vMerge continue 빈 셀 삽입
                    xml += '<w:tc>';
                    xml += '<w:tcPr>';
                    xml += `<w:tcW w:w="${tblGridWidths[colCursor] || 0}" w:type="dxa"/>`;
                    xml += '<w:vMerge/>';
                    xml += '</w:tcPr>';
                    xml += '<w:p/>';
                    xml += '</w:tc>';
                    colCursor++;
                } else if (cellIdx < cells.length) {
                    const tc = cells[cellIdx++];
                    // tc 직접 자식에서만 cellSpan 탐색 (중첩 표 혼입 방지)
                    const cellSpanElem = Array.from(tc.childNodes)
                        .find(n => n.nodeType === 1 && (n as Element).localName === 'cellSpan') as Element | undefined;
                    const colSpan = cellSpanElem
                        ? Math.max(1, parseInt(cellSpanElem.getAttribute('colSpan') || '1', 10))
                        : 1;
                    const rowSpan = cellSpanElem
                        ? Math.max(1, parseInt(cellSpanElem.getAttribute('rowSpan') || '1', 10))
                        : 1;

                    // rowSpan > 1 이면 아래 행들의 해당 열을 점유 등록
                    if (rowSpan > 1) {
                        for (let rs = 1; rs < rowSpan; rs++) {
                            for (let cs = 0; cs < colSpan; cs++) {
                                rowspanOccupied.add(`${r + rs},${colCursor + cs}`);
                            }
                        }
                    }

                    // tblGridWidths 기반 셀 폭 계산 (tblGrid와 일관성 유지)
                    let cellWidthDxa = 0;
                    for (let cs = 0; cs < colSpan; cs++) {
                        cellWidthDxa += tblGridWidths[colCursor + cs] || 0;
                    }

                    xml += this.convertTableCell(tc, colSpan, rowSpan, cellWidthDxa);
                    colCursor += colSpan;
                } else {
                    // 셀이 부족한 경우 → 빈 셀로 채워 DOCX 규칙 준수
                    xml += `<w:tc><w:tcPr><w:tcW w:w="${tblGridWidths[colCursor] || 0}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`;
                    colCursor++;
                }
            }

            xml += '</w:tr>';
        }

        xml += '</w:tbl>';
        return xml;
    }

    private convertTableCell(tcElem: Element, colSpan = 1, rowSpan = 1, cellWidthDxa = 0): string {
        let xml = '<w:tc>';
        xml += '<w:tcPr>';

        // 셀 너비: tblGridWidths 합산값(cellWidthDxa) 우선, 없으면 cellSz에서 파싱 (HWPUnit → dxa: ÷5)
        let tcWidth = cellWidthDxa;
        if (tcWidth === 0) {
            const cellSzElem = (
                Array.from(tcElem.childNodes).find(n => n.nodeType === 1 && (n as Element).localName === 'cellSz') ||
                Array.from(tcElem.childNodes).find(n => n.nodeType === 1 && (n as Element).localName === 'sz')
            ) as Element | undefined;
            const cellWidthHwp = cellSzElem ? parseInt(cellSzElem.getAttribute('width') || '0', 10) : 0;
            tcWidth = Math.round(cellWidthHwp / 5);
        }
        xml += `<w:tcW w:w="${tcWidth}" w:type="dxa"/>`;

        // colspan / rowspan
        if (colSpan > 1) xml += `<w:gridSpan w:val="${colSpan}"/>`;
        if (rowSpan > 1) xml += '<w:vMerge w:val="restart"/>';

        // 셀 테두리: tc 의 borderFillIDRef → borderFill 맵 조회
        const tcBfId = tcElem.getAttribute('borderFillIDRef') || '';
        const tcBf = this.ctx.borderFills.get(tcBfId);
        xml += this.buildTcBordersXml(tcBf);

        // 셀 배경색: borderFill.faceColor 가 있으면 <w:shd>
        if (tcBf?.faceColor) {
            const fill = tcBf.faceColor.replace('#', '');
            xml += `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`;
        }

        // 셀 여백: tc 직접 자식 <hp:cellMargin> (HWPUnit → dxa: ÷5)
        const cellMarginElem = Array.from(tcElem.childNodes)
            .find(n => n.nodeType === 1 && (n as Element).localName === 'cellMargin') as Element | undefined;
        if (cellMarginElem) {
            const toM = (attr: string) => Math.round(parseInt(cellMarginElem.getAttribute(attr) || '0', 10) / 5);
            const mL = toM('left'), mR = toM('right'), mT = toM('top'), mB = toM('bottom');
            xml += `<w:tcMar>`;
            xml += `<w:top w:w="${mT}" w:type="dxa"/>`;
            xml += `<w:left w:w="${mL}" w:type="dxa"/>`;
            xml += `<w:bottom w:w="${mB}" w:type="dxa"/>`;
            xml += `<w:right w:w="${mR}" w:type="dxa"/>`;
            xml += `</w:tcMar>`;
        }

        // 수직 정렬: subList.vertAlign 속성 우선, 없으면 center
        const subListElem = Array.from(tcElem.childNodes)
            .find(n => n.nodeType === 1 && (n as Element).localName === 'subList') as Element | undefined;
        const vAlignHwp = subListElem?.getAttribute('vertAlign') || 'CENTER';
        const vAlignMap: { [k: string]: string } = { 'TOP': 'top', 'CENTER': 'center', 'BOTTOM': 'bottom' };
        xml += `<w:vAlign w:val="${vAlignMap[vAlignHwp] || 'center'}"/>`;

        xml += '</w:tcPr>';

        // 셀 내용: <hp:subList>의 직접 자식 <hp:p>만 변환
        const subList = tcElem.getElementsByTagNameNS(this.ctx.ns.hp, 'subList')[0];
        let hasParagraph = false;
        if (subList) {
            const paragraphs = Array.from(subList.childNodes)
                .filter(n => n.nodeType === 1 && (n as Element).localName === 'p') as Element[];
            for (const p of paragraphs) {
                xml += this.convertParagraph(p);
                hasParagraph = true;
            }
        }

        // DOCX 규칙: <w:tc>는 반드시 하나 이상의 <w:p>를 가져야 함
        if (!hasParagraph) xml += '<w:p/>';

        xml += '</w:tc>';
        return xml;
    }

    private convertImage(picElem: Element, charPrId: string): string {
        const imgElem = picElem.getElementsByTagNameNS(this.ctx.ns.hc, 'img')[0];
        if (!imgElem) return '';

        const binaryItemIDRef = imgElem.getAttribute('binaryItemIDRef');
        if (!binaryItemIDRef || !this.ctx.images.has(binaryItemIDRef)) return '';

        const imgInfo = this.ctx.images.get(binaryItemIDRef);
        if (!imgInfo) return '';

        const szElem = picElem.getElementsByTagNameNS(this.ctx.ns.hp, 'sz')[0];
        // HWPUnit → EMU: 1 HWPUnit = 1/7200 inch, 1 inch = 914400 EMU → 914400/7200 = 127 EMU/HWPUnit
        // 폴백: 페이지 컨텐츠 폭의 50%를 너비로, 4:3 비율을 높이로 사용
        const contentWidthEmu = (this.ctx.pageWidth - this.ctx.marginLeft - this.ctx.marginRight) * 635; // twips→EMU (×635)
        let cx = Math.round(contentWidthEmu / 2);
        let cy = Math.round(cx * 3 / 4);
        if (szElem) {
            const w = parseInt(szElem.getAttribute('width') || '0');
            const h = parseInt(szElem.getAttribute('height') || '0');
            if (w > 0) cx = w * 127;
            if (h > 0) cy = h * 127;
        }

        const relId = `rId${++this.ctx.relIdCounter}`;
        const imgDocId = this.ctx.relIdCounter; // 고유 docPr id
        const fileName = `image${imgDocId}.${imgInfo.ext}`; // 고유 파일명 보장
        this.ctx.imageRels.push({
            id: relId,
            target: `media/${fileName}`,
            data: imgInfo.data,
            mediaType: imgInfo.mediaType,
            ext: imgInfo.ext
        });

        let xml = '<w:r>';
        xml += this.getRunPropertiesXml(charPrId);
        xml += '<w:drawing>';
        xml += `<wp:inline distT="0" distB="0" distL="0" distR="0">`;
        xml += `<wp:extent cx="${cx}" cy="${cy}"/>`;
        xml += `<wp:docPr id="${imgDocId}" name="Picture ${imgDocId}"/>`;
        xml += '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
        xml += '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">';
        xml += '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">';
        xml += `<pic:nvPicPr><pic:cNvPr id="${imgDocId}" name="Picture ${imgDocId}"/><pic:cNvPicPr/></pic:nvPicPr>`;
        xml += `<pic:blipFill><a:blip r:embed="${relId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`;
        xml += `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`;
        xml += '</pic:pic>';
        xml += '</a:graphicData>';
        xml += '</a:graphic>';
        xml += '</wp:inline>';
        xml += '</w:drawing>';
        xml += '</w:r>';
        return xml;
    }

    private convertRect(rectElem: Element, charPrId: string): string {
        let xml = '';
        const drawText = rectElem.getElementsByTagNameNS(this.ctx.ns.hp, 'drawText')[0];
        if (!drawText) return '';

        const subList = drawText.getElementsByTagNameNS(this.ctx.ns.hp, 'subList')[0];
        if (!subList) return '';

        const paragraphs = subList.getElementsByTagNameNS(this.ctx.ns.hp, 'p');
        for (const p of Array.from(paragraphs)) {
            if (p.parentNode !== subList) continue;
            const runs = p.getElementsByTagNameNS(this.ctx.ns.hp, 'run');
            for (const run of Array.from(runs)) {
                if (run.parentNode !== p) continue;
                const cId = run.getAttribute('charPrIDRef') || charPrId;
                const tElems = run.getElementsByTagNameNS(this.ctx.ns.hp, 't');
                for (const tElem of Array.from(tElems)) {
                    const text = this.getTextContent(tElem);
                    if (text) {
                        xml += '<w:r>';
                        xml += this.getRunPropertiesXml(cId);
                        xml += `<w:t xml:space="preserve">${this.escapeXml(text)}</w:t>`;
                        xml += '</w:r>';
                    }
                }
            }
        }
        return xml;
    }

    // ─── 표 테두리 보조 함수 ───────────────────────────────────────

    /**
     * HWPX border type → DOCX w:val 변환
     * HWPX width(mm 문자열) → DOCX sz (1/8 포인트 단위)
     */
    private hwpBorderToDocx(type: string, widthStr: string, color: string)
        : { val: string; sz: number; color: string } {
        const typeMap: { [k: string]: string } = {
            'NONE': 'none', 'SOLID': 'single', 'DOUBLE': 'double',
            'DASHED': 'dashed', 'DOTTED': 'dotted', 'DOUBLE_THIN': 'double'
        };
        const val = typeMap[type] || 'none';
        // widthStr 예: "0.12 mm", "0.1 mm"
        const mm = parseFloat(widthStr);
        const sz = isNaN(mm) ? 2 : Math.max(2, Math.round(mm * 72 / 25.4 * 8));
        const col = color ? color.replace('#', '') : '000000';
        return { val, sz, color: col };
    }

    /** tbl 레벨 <w:tblBorders> 생성 */
    private buildTblBordersXml(bf: import('./types').BorderFill | undefined): string {
        const sides = ['top', 'left', 'bottom', 'right'] as const;
        const bfMap: { [k: string]: import('./types').Border | undefined } = {
            top: bf?.topBorder, left: bf?.leftBorder,
            bottom: bf?.bottomBorder, right: bf?.rightBorder
        };

        let xml = '<w:tblBorders>';
        for (const side of sides) {
            const b = bfMap[side];
            if (b && b.type !== 'NONE') {
                const { val, sz, color: col } = this.hwpBorderToDocx(b.type, b.width, b.color);
                xml += `<w:${side} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${col}"/>`;
            } else {
                // 테두리 정보가 없거나 NONE → nil (상위 스타일 상속 차단)
                xml += `<w:${side} w:val="nil"/>`;
            }
        }
        // insideH / insideV: 외곽 테두리 중 유효한 것이 있으면 동일 스타일 적용
        const validBorders = bf
            ? [bf.topBorder, bf.leftBorder, bf.bottomBorder, bf.rightBorder].filter(b => b && b.type !== 'NONE')
            : [];
        if (validBorders.length > 0) {
            const rep = validBorders[0]!;
            const { val, sz, color: col } = this.hwpBorderToDocx(rep.type, rep.width, rep.color);
            xml += `<w:insideH w:val="${val}" w:sz="${sz}" w:space="0" w:color="${col}"/>`;
            xml += `<w:insideV w:val="${val}" w:sz="${sz}" w:space="0" w:color="${col}"/>`;
        } else {
            xml += '<w:insideH w:val="nil"/>';
            xml += '<w:insideV w:val="nil"/>';
        }
        xml += '</w:tblBorders>';
        return xml;
    }

    /** tc 레벨 <w:tcBorders> 생성 */
    private buildTcBordersXml(bf: import('./types').BorderFill | undefined): string {
        if (!bf) return '';
        const sides: { key: 'topBorder' | 'leftBorder' | 'bottomBorder' | 'rightBorder'; name: string }[] = [
            { key: 'topBorder', name: 'top' },
            { key: 'leftBorder', name: 'left' },
            { key: 'bottomBorder', name: 'bottom' },
            { key: 'rightBorder', name: 'right' }
        ];
        let xml = '<w:tcBorders>';
        for (const { key, name } of sides) {
            const b = bf[key];
            if (b) {
                const { val, sz, color: col } = this.hwpBorderToDocx(b.type, b.width, b.color);
                xml += `<w:${name} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${col}"/>`;
            }
        }
        xml += '</w:tcBorders>';
        return xml;
    }

    private escapeXml(text: string): string {
        if (!text) return '';
        return text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }
}
