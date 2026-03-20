import { ConversionState } from './state';
import { escapeXml, getFontFamilyType } from './utils';
import { BorderDef } from './types';

export class HwpxGenerator {
    private state: ConversionState;

    constructor(state: ConversionState) {
        this.state = state;
    }

    private createFontFaces() {
        const langs = ['HANGUL', 'LATIN', 'HANJA', 'JAPANESE', 'OTHER', 'SYMBOL', 'USER'];
        let xml = `<hh:fontfaces itemCnt="${langs.length}">`;

        for (const lang of langs) {
            const fontMap = this.state.langFontFaces[lang];
            xml += `<hh:fontface lang="${lang}" fontCnt="${fontMap.size}">`;
            for (const [face, id] of fontMap) {
                const familyType = getFontFamilyType(face);
                xml += `<hh:font id="${id}" face="${face}" type="TTF" isEmbedded="0">`;
                xml += `<hh:typeInfo familyType="${familyType}" weight="0" proportion="0" contrast="0" strokeVariation="0" armStyle="0" letterform="0" midline="0" xHeight="0"/>`;
                xml += `</hh:font>`;
            }
            xml += `</hh:fontface>`;
        }
        xml += `</hh:fontfaces>`;
        return xml;
    }



    private createSection(doc: Document) {
        let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec xmlns:ha="${this.state.hwpxNs.ha}" xmlns:hp="${this.state.hwpxNs.hp}" xmlns:hp10="${this.state.hwpxNs.hp10}" xmlns:hs="${this.state.hwpxNs.hs}" xmlns:hc="${this.state.hwpxNs.hc}" xmlns:hh="${this.state.hwpxNs.hh}" xmlns:hhs="${this.state.hwpxNs.hhs}" xmlns:hm="${this.state.hwpxNs.hm}" xmlns:hpf="${this.state.hwpxNs.hpf}" xmlns:dc="${this.state.hwpxNs.dc}" xmlns:opf="${this.state.hwpxNs.opf}" xmlns:ooxmlchart="${this.state.hwpxNs.ooxmlchart}" xmlns:epub="${this.state.hwpxNs.epub}" xmlns:config="${this.state.hwpxNs.config}">`;

        const body = doc.getElementsByTagNameNS(this.state.ns.w, 'body')[0];
        if (body) {
            const bodyChildren = Array.from(body.children).filter(
                c => c.localName === 'p' || c.localName === 'tbl' || c.localName === 'sdt'
            );
            const totalChildren = bodyChildren.length;

            bodyChildren.forEach((child, index) => {
                const isLastElement = (index === totalChildren - 1);
                if (child.localName === 'p') {
                    xml += this.convertParagraph(child, isLastElement);
                } else if (child.localName === 'tbl') {
                    xml += this.convertTable(child);
                } else if (child.localName === 'sdt') {
                    const sdtContent = child.getElementsByTagNameNS(this.state.ns.w, 'sdtContent')[0];
                    if (sdtContent) {
                        const sdtChildren = Array.from(sdtContent.children).filter(
                            c => c.localName === 'p' || c.localName === 'tbl'
                        );
                        const totalSdtChildren = sdtChildren.length;
                        sdtChildren.forEach((sdtChild, sdtIndex) => {
                            const isLastSdtElement = isLastElement && (sdtIndex === totalSdtChildren - 1);
                            if (sdtChild.localName === 'p') {
                                xml += this.convertParagraph(sdtChild, isLastSdtElement);
                            } else if (sdtChild.localName === 'tbl') {
                                xml += this.convertTable(sdtChild);
                            }
                        });
                    }
                }
            });
        } else {
            const paraId = this.state.idCounter++;
            xml += `<hp:p id="${paraId}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="2">${this.createSecPr()}<hp:t></hp:t></hp:run><hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1265" textheight="1265" baseline="1032" spacing="188" horzpos="0" horzsize="${this.state.TEXT_WIDTH}" flags="393216"/></hp:linesegarray></hp:p>`;
        }
        xml += `</hs:sec>`;
        return xml;
    }

    private createSecPr() {
        return `<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="${this.state.COL_GAP}" tabStop="8000" outlineShapeIDRef="1" memoShapeIDRef="1" textVerticalWidthHead="0" masterPageCnt="0"><hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/><hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/><hp:pagePr landscape="WIDELY" width="${this.state.PAGE_WIDTH}" height="${this.state.PAGE_HEIGHT}" gutterType="LEFT_ONLY"><hp:margin header="${this.state.HEADER_MARGIN}" footer="${this.state.FOOTER_MARGIN}" gutter="${this.state.GUTTER_MARGIN}" left="${this.state.MARGIN_LEFT}" right="${this.state.MARGIN_RIGHT}" top="${this.state.MARGIN_TOP}" bottom="${this.state.MARGIN_BOTTOM}"/></hp:pagePr><hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/><hp:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/><hp:noteSpacing betweenNotes="283" belowLine="0" aboveLine="1000"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr><hp:endNotePr><hp:autoNumFormat type="ROMAN_SMALL" userChar="" prefixChar="" suffixChar="" supscript="1"/><hp:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="0" aboveLine="1000"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr><hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill></hp:secPr><hp:ctrl><hp:colPr id="" type="${this.state.COL_TYPE}" layout="LEFT" colCount="${this.state.COL_COUNT}" sameSz="1" sameGap="0"/></hp:ctrl>`;
    }

    private convertParagraph(para: Element, _isLastElement = false) {
        const paraId = this.state.idCounter++;
        const paraPrId = this.getParaShapeId(para, _isLastElement);
        let paraStyleId: string | null = null;

        const pPr = para.getElementsByTagNameNS(this.state.ns.w, 'pPr')[0];
        if (pPr) {
            const pStyle = pPr.getElementsByTagNameNS(this.state.ns.w, 'pStyle')[0];
            if (pStyle) paraStyleId = pStyle.getAttribute('w:val');
        }

        // JS 원본과 동일하게 실제 폰트 크기 계산
        const baseFontSize = this.getParagraphFontSize(para);
        const fontSizeInHWPUNIT = Math.round(baseFontSize * this.state.HWPUNIT_PER_INCH / 72);

        const paraShape = this.state.paraShapes.get(paraPrId);
        // JS 원본: 기본 lineSpacing은 115
        const lineSpacingPercent = (paraShape && paraShape.lineSpacingType === 'PERCENT') ? paraShape.lineSpacingVal : 115;
        const textheight = fontSizeInHWPUNIT;
        const baseline = Math.round(textheight * 0.85);
        const spacing = Math.round(textheight * (lineSpacingPercent - 100) / 100);
        const vertsize = textheight + spacing;

        let pageBreak = '0';
        // JS 원본과 동일: pPr.rPr에서 처리
        if (pPr) {
            const pageBreakBefore = pPr.getElementsByTagNameNS(this.state.ns.w, 'pageBreakBefore')[0];
            if (pageBreakBefore && pageBreakBefore.getAttribute('w:val') !== 'false') pageBreak = '1';
            const rPr = pPr.getElementsByTagNameNS(this.state.ns.w, 'rPr')[0];
            this.getCharShapeId(rPr || null, baseFontSize, false, paraStyleId);
        } else {
            this.getCharShapeId(null, baseFontSize, false, paraStyleId);
        }

        let hwpxStyleId = '0';
        if (paraStyleId && this.state.docxStyleToHwpxId && this.state.docxStyleToHwpxId[paraStyleId] !== undefined) {
            hwpxStyleId = String(this.state.docxStyleToHwpxId[paraStyleId]);
        }

        let xml = `<hp:p id="${paraId}" paraPrIDRef="${paraPrId}" styleIDRef="${hwpxStyleId}" pageBreak="${pageBreak}" columnBreak="0" merged="0">`;

        if (this.state.isFirstParagraph) {
            xml += `<hp:run charPrIDRef="0">${this.createSecPr()}</hp:run>`;
            this.state.isFirstParagraph = false;
        }

        // 목록 카운터 추적
        const numPr = pPr?.getElementsByTagNameNS(this.state.ns.w, 'numPr')[0];
        if (numPr) {
            const numId = numPr.getElementsByTagNameNS(this.state.ns.w, 'numId')[0]?.getAttribute('w:val');
            const ilvl = numPr.getElementsByTagNameNS(this.state.ns.w, 'ilvl')[0]?.getAttribute('w:val');
            if (numId && ilvl) this.getListPrefix(numId, ilvl);
        }

        const runs = para.childNodes;
        for (let i = 0; i < runs.length; i++) {
            const node = runs[i] as Element;
            if (node.nodeName === 'w:r') {
                xml += this.convertRun(node, baseFontSize, paraStyleId);
            } else if (node.nodeName === 'w:hyperlink') {
                for (let j = 0; j < node.childNodes.length; j++) {
                    const child = node.childNodes[j] as Element;
                    if (child.nodeName === 'w:r') {
                        xml += this.convertRun(child, baseFontSize, paraStyleId);
                    }
                }
            }
        }

        const actualPrevSpacing = (paraShape && paraShape.prevSpacing) ? paraShape.prevSpacing : 0;
        const actualNextSpacing = (paraShape && paraShape.nextSpacing) ? paraShape.nextSpacing : 0;

        this.state.currentVertPos += actualPrevSpacing;
        xml += `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${this.state.currentVertPos}" vertsize="${vertsize}" textheight="${textheight}" baseline="${baseline}" spacing="${spacing}" horzpos="0" horzsize="${this.state.TEXT_WIDTH}" flags="393216"/></hp:linesegarray></hp:p>`;
        this.state.currentVertPos += vertsize + actualNextSpacing;

        return xml;
    }

    private convertRun(run: Element, baseFontSize: number | null, paraStyleId: string | null = null) {
        const rPr = run.getElementsByTagNameNS(this.state.ns.w, 'rPr')[0];
        let isHyperlink = false;
        if (run.parentNode && run.parentNode.nodeName === 'w:hyperlink') {
            isHyperlink = true;
        }

        const charPrId = this.getCharShapeId(rPr, baseFontSize, isHyperlink, paraStyleId);

        let xml = '';
        let textBuffer = '';
        let hasContentInRun = false;

        const flushTextBuffer = () => {
            if (textBuffer.length > 0 || hasContentInRun) {
                xml += `<hp:run charPrIDRef="${charPrId}">`;
                xml += textBuffer;
                xml += `</hp:run>`;
                textBuffer = '';
                hasContentInRun = false;
            }
        };

        for (let i = 0; i < run.childNodes.length; i++) {
            const child = run.childNodes[i] as Element;
            if (child.nodeName === 'w:t') {
                textBuffer += `<hp:t>${escapeXml(child.textContent || '')}</hp:t>`;
                hasContentInRun = true;
            } else if (child.nodeName === 'w:tab') {
                textBuffer += `<hp:t>\t</hp:t>`;
                hasContentInRun = true;
            } else if (child.nodeName === 'w:br') {
                if (child.getAttribute('w:type') !== 'page') {
                    textBuffer += `<hp:t>\n</hp:t>`;
                    hasContentInRun = true;
                }
            } else if (child.nodeName === 'w:footnoteReference') {
                const id = child.getAttribute('w:id');
                if (id) {
                    textBuffer += this.convertFootnote(id, 'FOOTNOTE');
                    hasContentInRun = true;
                }
            } else if (child.nodeName === 'w:endnoteReference') {
                const id = child.getAttribute('w:id');
                if (id) {
                    textBuffer += this.convertFootnote(id, 'ENDNOTE');
                    hasContentInRun = true;
                }
            } else if (child.nodeName === 'w:drawing') {
                flushTextBuffer();
                xml += this.convertDrawing(child);
            } else if (child.localName === 'AlternateContent' || child.nodeName.includes('AlternateContent')) {
                const allElements = child.getElementsByTagName('*');
                const drawings = Array.from(allElements).filter(el => el.localName === 'drawing');

                if (drawings.length > 0) {
                    flushTextBuffer();
                    xml += this.convertDrawing(drawings[0]);
                }
            }
        }
        flushTextBuffer();
        return xml;
    }

    private convertFootnote(id: string, type: 'FOOTNOTE' | 'ENDNOTE') {
        const map = type === 'FOOTNOTE' ? this.state.footnoteMap : this.state.endnoteMap;
        const node = map.get(id);
        if (!node) return '';

        const isFootnote = type === 'FOOTNOTE';
        const number = isFootnote ? ++this.state.footnoteNumber : ++this.state.endnoteNumber;
        const instId = Math.floor(Math.random() * 10000000);

        const paras = node.getElementsByTagNameNS(this.state.ns.w, 'p');
        let contentXml = '';
        let isFirstPara = true;

        for (let i = 0; i < paras.length; i++) {
            const para = paras[i];
            if (isFirstPara) {
                const paraContent = this.convertParagraph(para);
                const autoNumType = isFootnote ? 'FOOTNOTE' : 'ENDNOTE';
                const autoNumXml = `<hp:ctrl><hp:autoNum num="${number}" numType="${autoNumType}"><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/></hp:autoNum></hp:ctrl>`;
                // First hp:run check might be brittle, but matches js logic
                contentXml += paraContent.replace(/(<hp:run[^>]*>)/, `$1${autoNumXml}`);
                isFirstPara = false;
            } else {
                contentXml += this.convertParagraph(para);
            }
        }

        const tagName = isFootnote ? 'hp:footNote' : 'hp:endNote';
        const flagAttr = isFootnote ? '' : ' flag="3"';
        return `<hp:ctrl><${tagName} number="${number}" instId="${instId}"${flagAttr}><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${contentXml}</hp:subList></${tagName}></hp:ctrl>`;
    }

    private getParaShapeId(para: Element | null, forceZeroNextSpacing = false) {
        if (!para || typeof para.getElementsByTagNameNS !== 'function') return '1';

        const pPr = para.getElementsByTagNameNS(this.state.ns.w, 'pPr')[0];
        let align = 'LEFT';
        let heading = 'NONE';
        let level = '0';
        let headingIdRef = '0';
        let leftMargin = 0;
        let rightMargin = 0;
        let indent = 0;
        let prevSpacing = 0;
        let nextSpacing = 0;
        let lineSpacingType = 'PERCENT';
        let lineSpacingVal = 115;
        let keepWithNext = '0';
        let keepLines = '0';
        let pageBreakBefore = '0';
        let borderFillIDRef = '1';
        let hasDirectInd = false;
        let hasDirectSpacing = false;
        let paraStyleId: string | null = null;

        if (pPr) {
            // 정렬
            let jcVal: string | null = null;
            const jc = pPr.getElementsByTagNameNS(this.state.ns.w, 'jc')[0];
            if (jc) {
                jcVal = jc.getAttribute('w:val');
            } else {
                const pStyle = pPr.getElementsByTagNameNS(this.state.ns.w, 'pStyle')[0];
                if (pStyle) {
                    paraStyleId = pStyle.getAttribute('w:val');
                    let currentStyleId: string | null | undefined = paraStyleId;
                    while (currentStyleId) {
                        const style = this.state.docStyles.get(currentStyleId);
                        if (!style) break;
                        if (style.align) { jcVal = style.align; break; }
                        currentStyleId = style.basedOn;
                    }
                }
            }
            if (jcVal === 'center') align = 'CENTER';
            else if (jcVal === 'right') align = 'RIGHT';
            else if (jcVal === 'both') align = 'JUSTIFY';
            else if (jcVal === 'distribute') align = 'DISTRIBUTE';

            // 들여쓰기
            const ind = pPr.getElementsByTagNameNS(this.state.ns.w, 'ind')[0];
            if (ind) {
                hasDirectInd = true;
                const left = ind.getAttribute('w:left') || ind.getAttribute('w:start');
                const right = ind.getAttribute('w:right') || ind.getAttribute('w:end');
                const firstLine = ind.getAttribute('w:firstLine');
                const hanging = ind.getAttribute('w:hanging');
                if (left) leftMargin = parseInt(left) * 5;
                if (right) rightMargin = parseInt(right) * 5;
                if (firstLine) indent = parseInt(firstLine) * 5;
                else if (hanging) indent = -parseInt(hanging) * 5;
            }

            // 목록(numPr)
            const numPr = pPr.getElementsByTagNameNS(this.state.ns.w, 'numPr')[0];
            if (numPr) {
                const numId = numPr.getElementsByTagNameNS(this.state.ns.w, 'numId')[0]?.getAttribute('w:val');
                const ilvl = numPr.getElementsByTagNameNS(this.state.ns.w, 'ilvl')[0]?.getAttribute('w:val') || '0';

                if (numId && numId !== '0') {
                    const absRef = this.state.numberingMap.get(numId);
                    if (absRef) {
                        let numberingIdx = 1;
                        for (const [key] of this.state.abstractNumMap) {
                            if (key === absRef) break;
                            numberingIdx++;
                        }

                        heading = 'NUMBER';
                        headingIdRef = String(numberingIdx);
                        level = String(parseInt(ilvl) + 1);

                        if (!hasDirectInd) {
                            const levels = this.state.abstractNumMap.get(absRef);
                            if (levels) {
                                const lvl = levels.get(ilvl);
                                if (lvl && lvl.lvlLeft) {
                                    leftMargin = parseInt(lvl.lvlLeft) * 5;
                                    if (lvl.lvlHanging) indent = -parseInt(lvl.lvlHanging) * 5;
                                    hasDirectInd = true;
                                } else {
                                    leftMargin = (parseInt(ilvl) + 1) * 800 * 5;
                                    indent = -400 * 5;
                                    hasDirectInd = true;
                                }
                            }
                        }
                    }
                } else if (numId === '0') {
                    heading = 'NONE';
                }
            }
            // 줄 간격
            const spacing = pPr.getElementsByTagNameNS(this.state.ns.w, 'spacing')[0];
            if (spacing) {
                hasDirectSpacing = true;
                const before = spacing.getAttribute('w:before');
                const after = spacing.getAttribute('w:after');
                if (before) prevSpacing = parseInt(before) * 5;
                if (after) nextSpacing = 0; // JS 원본과 동일하게 after는 0으로 처리
                const line = spacing.getAttribute('w:line');
                const lineRule = spacing.getAttribute('w:lineRule');
                if (line) {
                    if (lineRule === 'auto' || !lineRule) {
                        lineSpacingVal = Math.round(parseInt(line) / 240 * 100);
                    } else if (lineRule === 'exact') {
                        lineSpacingType = 'FIXED';
                        lineSpacingVal = parseInt(line) * 5;
                    } else if (lineRule === 'atLeast') {
                        lineSpacingType = 'ATLEAST';
                        lineSpacingVal = parseInt(line) * 5;
                    }
                }
            }
        }

        // 스타일 상속
        if ((!hasDirectInd || !hasDirectSpacing) && paraStyleId) {
            let currentStyleId: string | null | undefined = paraStyleId;
            while (currentStyleId) {
                const style = this.state.docStyles.get(currentStyleId);
                if (!style) break;
                if (!hasDirectInd && style.ind) {
                    if (style.ind.left) leftMargin = parseInt(style.ind.left) * 5;
                    if (style.ind.right) rightMargin = parseInt(style.ind.right) * 5;
                    if (style.ind.firstLine) indent = parseInt(style.ind.firstLine) * 5;
                    if (style.ind.hanging) indent = -parseInt(style.ind.hanging) * 5;
                    hasDirectInd = true;
                }
                if (!hasDirectSpacing && style.spacing) {
                    if (style.spacing.before) prevSpacing = parseInt(style.spacing.before) * 5;
                    if (style.spacing.after) nextSpacing = parseInt(style.spacing.after) * 5;
                    if (style.spacing.line) {
                        const rule = style.spacing.lineRule;
                        if (rule === 'auto' || !rule) {
                            lineSpacingVal = Math.round(parseInt(style.spacing.line) / 240 * 100);
                        } else if (rule === 'exact') {
                            lineSpacingType = 'FIXED';
                            lineSpacingVal = parseInt(style.spacing.line) * 5;
                        } else if (rule === 'atLeast') {
                            lineSpacingType = 'ATLEAST';
                            lineSpacingVal = parseInt(style.spacing.line) * 5;
                        }
                    }
                    hasDirectSpacing = true;
                }
                if (hasDirectInd && hasDirectSpacing) break;
                currentStyleId = style.basedOn;
            }
        }

        if (pPr) {
            // 아웃라인
            const outlineLvl = pPr.getElementsByTagNameNS(this.state.ns.w, 'outlineLvl')[0];
            if (outlineLvl) {
                //heading = 'OUTLINE';
                level = outlineLvl.getAttribute('w:val') || '0';
            }

            // keepWithNext
            const kwn = pPr.getElementsByTagNameNS(this.state.ns.w, 'keepNext')[0];
            if (kwn && kwn.getAttribute('w:val') !== 'false' && kwn.getAttribute('w:val') !== '0') keepWithNext = '1';

            // keepLines
            const kl = pPr.getElementsByTagNameNS(this.state.ns.w, 'keepLines')[0];
            if (kl && kl.getAttribute('w:val') !== 'false' && kl.getAttribute('w:val') !== '0') keepLines = '1';

            // pageBreakBefore
            const pbb = pPr.getElementsByTagNameNS(this.state.ns.w, 'pageBreakBefore')[0];
            if (pbb && pbb.getAttribute('w:val') !== 'false' && pbb.getAttribute('w:val') !== '0') pageBreakBefore = '1';

            // 테두리/배경 (pBdr / shd)
            const pBdr = pPr.getElementsByTagNameNS(this.state.ns.w, 'pBdr')[0];
            const shd = pPr.getElementsByTagNameNS(this.state.ns.w, 'shd')[0];

            if (pBdr || shd) {
                const parseBorder = (node: Element | undefined): BorderDef => {
                    if (!node) return { type: 'NONE', width: '0.1 mm', color: '#000000' };
                    const val = node.getAttribute('w:val');
                    if (val === 'none' || val === 'nil') return { type: 'NONE', width: '0.1 mm', color: '#000000' };
                    const sz = node.getAttribute('w:sz');
                    let width = '0.1 mm';
                    if (sz) width = (parseInt(sz) / 8 * 0.3528).toFixed(2) + ' mm';
                    const colorAttr = node.getAttribute('w:color');
                    let color = '#000000';
                    if (colorAttr && colorAttr !== 'auto') color = `#${colorAttr}`;
                    let type = 'SOLID';
                    if (val === 'double') type = 'DOUBLE';
                    else if (val === 'dotted') type = 'DOT';
                    else if (val === 'dashed') type = 'DASH';
                    return { type, width, color };
                };
                const left = parseBorder(pBdr?.getElementsByTagNameNS(this.state.ns.w, 'left')[0]);
                const right = parseBorder(pBdr?.getElementsByTagNameNS(this.state.ns.w, 'right')[0]);
                const top = parseBorder(pBdr?.getElementsByTagNameNS(this.state.ns.w, 'top')[0]);
                const bottom = parseBorder(pBdr?.getElementsByTagNameNS(this.state.ns.w, 'bottom')[0]);
                let backColor = '#FFFFFF';
                if (shd) {
                    const fill = shd.getAttribute('w:fill');
                    if (fill && fill !== 'auto') backColor = `#${fill}`;
                }
                const bfKey = `${left.type}_${left.width}_${left.color}|${right.type}_${right.width}_${right.color}|${top.type}_${top.width}_${top.color}|${bottom.type}_${bottom.width}_${bottom.color}|${backColor}`;
                let foundId: string | null = null;
                for (const [id, bf] of this.state.borderFills) {
                    const check = `${bf.leftBorder.type}_${bf.leftBorder.width}_${bf.leftBorder.color}|${bf.rightBorder.type}_${bf.rightBorder.width}_${bf.rightBorder.color}|${bf.topBorder.type}_${bf.topBorder.width}_${bf.topBorder.color}|${bf.bottomBorder.type}_${bf.bottomBorder.width}_${bf.bottomBorder.color}|${bf.backColor || '#FFFFFF'}`;
                    if (bfKey === check) { foundId = id; break; }
                }
                if (foundId) {
                    borderFillIDRef = foundId;
                } else {
                    const newBfId = String(this.state.borderFillIdCounter++);
                    this.state.borderFills.set(newBfId, {
                        id: newBfId, leftBorder: left, rightBorder: right, topBorder: top, bottomBorder: bottom,
                        diagonal: null, slash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
                        backSlash: { type: 'NONE', width: '0.1 mm', color: '#000000' }, backColor
                    });
                    borderFillIDRef = newBfId;
                }
            }
        }

        if (forceZeroNextSpacing) nextSpacing = 0;

        const key = `${align}_${leftMargin}_${rightMargin}_${indent}_${prevSpacing}_${nextSpacing}_${lineSpacingType}_${lineSpacingVal}_${heading}_${level}_${headingIdRef}_${borderFillIDRef}_${keepWithNext}_${keepLines}_${pageBreakBefore}`;

        for (const [id, ps] of this.state.paraShapes) {
            if (ps._key === key) return id;
        }

        const newId = String(this.state.paraShapes.size);
        this.state.paraShapes.set(newId, {
            id: newId, align, heading, level, headingIdRef,
            leftMargin, rightMargin, indent, prevSpacing, nextSpacing,
            lineSpacingType, lineSpacingVal, borderFillIDRef,
            keepWithNext, keepLines, pageBreakBefore,
            tabPrIDRef: '1', tabStops: [], _key: key
        });
        return newId;
    }

    private getCharShapeId(rPr: Element | null, baseFontSize: number | null, isHyperlink = false, paraStyleId: string | null = null) {
        let bold = false;
        let italic = false;
        let underline = isHyperlink;
        let strike = false;
        let supscript = false;
        let subscript = false;
        let color = isHyperlink ? '#0000FF' : '#000000';
        let shadeColor = 'none';
        let borderFillIDRef = '1';

        let fontSize = baseFontSize ? Math.round(baseFontSize * 100) : 1000;
        let directRFonts: Element | null = null;
        let hasDirectSize = false;

        if (rPr) {
            if (rPr.getElementsByTagNameNS(this.state.ns.w, 'b').length > 0 && rPr.getElementsByTagNameNS(this.state.ns.w, 'b')[0].getAttribute('w:val') !== '0') bold = true;
            if (rPr.getElementsByTagNameNS(this.state.ns.w, 'i').length > 0 && rPr.getElementsByTagNameNS(this.state.ns.w, 'i')[0].getAttribute('w:val') !== '0') italic = true;
            if (rPr.getElementsByTagNameNS(this.state.ns.w, 'u').length > 0 && rPr.getElementsByTagNameNS(this.state.ns.w, 'u')[0].getAttribute('w:val') !== 'none') underline = true;
            if (rPr.getElementsByTagNameNS(this.state.ns.w, 'strike').length > 0 && rPr.getElementsByTagNameNS(this.state.ns.w, 'strike')[0].getAttribute('w:val') !== '0') strike = true;
            if (rPr.getElementsByTagNameNS(this.state.ns.w, 'dstrike').length > 0 && rPr.getElementsByTagNameNS(this.state.ns.w, 'dstrike')[0].getAttribute('w:val') !== '0') strike = true;

            const vertAlign = rPr.getElementsByTagNameNS(this.state.ns.w, 'vertAlign')[0];
            if (vertAlign) {
                const val = vertAlign.getAttribute('w:val');
                if (val === 'superscript') supscript = true;
                else if (val === 'subscript') subscript = true;
            }

            const colorElem = rPr.getElementsByTagNameNS(this.state.ns.w, 'color')[0];
            if (colorElem) {
                const val = colorElem.getAttribute('w:val');
                if (val && val !== 'auto') color = `#${val}`;
            }

            const sz = rPr.getElementsByTagNameNS(this.state.ns.w, 'sz')[0];
            if (sz) {
                const val = parseInt(sz.getAttribute('w:val') || '0');
                if (!isNaN(val) && val > 0) { fontSize = val * 50; hasDirectSize = true; }
            }

            // Highlight
            const highlight = rPr.getElementsByTagNameNS(this.state.ns.w, 'highlight')[0];
            if (highlight) {
                const val = highlight.getAttribute('w:val') || '';
                const highlightMap: { [key: string]: string } = {
                    'black': '#000000', 'blue': '#0000FF', 'cyan': '#00FFFF', 'green': '#008000',
                    'magenta': '#FF00FF', 'red': '#FF0000', 'yellow': '#FFFF00', 'white': '#FFFFFF',
                    'darkBlue': '#00008B', 'darkCyan': '#008B8B', 'darkGreen': '#006400',
                    'darkMagenta': '#8B008B', 'darkRed': '#8B0000', 'darkYellow': '#808000',
                    'darkGray': '#A9A9A9', 'lightGray': '#D3D3D3', 'none': 'none'
                };
                if (highlightMap[val] && highlightMap[val] !== 'none') {
                    shadeColor = highlightMap[val];
                    borderFillIDRef = this.getOrCreateHighlightBorderFill(shadeColor);
                }
            }

            // Shading
            const shd = rPr.getElementsByTagNameNS(this.state.ns.w, 'shd')[0];
            if (shd) {
                const fill = shd.getAttribute('w:fill');
                if (fill && fill !== 'auto') {
                    if (fill === '000000' || fill === 'black') {
                        borderFillIDRef = '6';
                        if (color === '#000000') color = '#FFFFFF';
                    } else {
                        shadeColor = `#${fill}`;
                        borderFillIDRef = this.getOrCreateHighlightBorderFill(shadeColor);
                    }
                }
            }

            // Border (Box Text)
            const bdr = rPr.getElementsByTagNameNS(this.state.ns.w, 'bdr')[0];
            if (bdr) borderFillIDRef = '5';

            // Font
            const rFonts = rPr.getElementsByTagNameNS(this.state.ns.w, 'rFonts')[0];
            if (rFonts) {
                directRFonts = rFonts;
            }
        }

        // 스타일 상속: 폰트 크기
        if (!hasDirectSize && paraStyleId) {
            let currentStyleId: string | null | undefined = paraStyleId;
            while (currentStyleId) {
                const style = this.state.docStyles.get(currentStyleId);
                if (!style) break;
                if (style.rPr && style.rPr.sz) {
                    const val = parseInt(style.rPr.sz);
                    if (!isNaN(val)) { fontSize = val * 50; break; }
                }
                currentStyleId = style.basedOn;
            }
        }

        // 스타일 상속: 글자 색상
        if (color === '#000000' && !isHyperlink && paraStyleId) {
            let currentStyleId: string | null | undefined = paraStyleId;
            while (currentStyleId) {
                const style = this.state.docStyles.get(currentStyleId);
                if (!style) break;
                if (style.rPr && style.rPr.color && style.rPr.color !== 'auto') {
                    color = `#${style.rPr.color}`; break;
                }
                currentStyleId = style.basedOn;
            }
        }

        const fontIds = this.registerFontsFromRFonts(directRFonts);
        const height = fontSize;

        const key = `${bold}_${italic}_${underline}_${strike}_${supscript}_${subscript}_${color}_${shadeColor}_${borderFillIDRef}_${height}_${fontIds.hangulId}_${fontIds.latinId}`;

        for (const [id, cs] of this.state.charShapes) {
            if (cs._key === key) return id;
        }

        const newId = String(this.state.charShapes.size);
        this.state.charShapes.set(newId, {
            id: newId,
            height: String(height),
            textColor: color,
            shadeColor,
            borderFillIDRef,
            bold, italic, underline, strike, supscript, subscript,
            fontIds,
            fontId: fontIds.hangulId,
            _key: key
        });
        return newId;
    }

    private getOrCreateHighlightBorderFill(highlightColor: string): string {
        for (const [id, bf] of this.state.borderFills) {
            if (bf.backColor === highlightColor &&
                bf.leftBorder.type === 'NONE' && bf.rightBorder.type === 'NONE' &&
                bf.topBorder.type === 'NONE' && bf.bottomBorder.type === 'NONE') {
                return id;
            }
        }
        const newBfId = String(this.state.borderFillIdCounter++);
        this.state.borderFills.set(newBfId, {
            id: newBfId,
            leftBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            rightBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            topBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            bottomBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            diagonal: null,
            slash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backSlash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backColor: highlightColor
        });
        return newBfId;
    }

    private registerFontsFromRFonts(rFonts: Element | null): { hangulId: number, latinId: number, hanjaId: number, japaneseId: number, otherId: number, symbolId: number, userId: number } {
        let hangulFont = '함초롬바탕';
        let latinFont = '함초롬바탕';
        if (rFonts) {
            const eastAsia = rFonts.getAttribute('w:eastAsia');
            const ascii = rFonts.getAttribute('w:ascii');
            const hAnsi = rFonts.getAttribute('w:hAnsi');
            if (eastAsia) hangulFont = eastAsia;
            if (ascii) latinFont = ascii;
            else if (hAnsi) latinFont = hAnsi;
        }
        return {
            hangulId: this.state.registerFontForLang('HANGUL', hangulFont),
            latinId: this.state.registerFontForLang('LATIN', latinFont),
            hanjaId: this.state.registerFontForLang('HANJA', hangulFont),
            japaneseId: this.state.registerFontForLang('JAPANESE', hangulFont),
            otherId: this.state.registerFontForLang('OTHER', latinFont),
            symbolId: this.state.registerFontForLang('SYMBOL', latinFont),
            userId: this.state.registerFontForLang('USER', hangulFont),
        };
    }


    private getParagraphFontSize(para: Element): number {
        const runs = para.getElementsByTagNameNS(this.state.ns.w, 'r');
        for (let i = 0; i < runs.length; i++) {
            const rPr = runs[i].getElementsByTagNameNS(this.state.ns.w, 'rPr')[0];
            if (rPr) {
                const sz = rPr.getElementsByTagNameNS(this.state.ns.w, 'sz')[0];
                if (sz) {
                    const val = sz.getAttribute('w:val');
                    if (val) return parseInt(val) / 2;
                }
            }
        }
        const pPr = para.getElementsByTagNameNS(this.state.ns.w, 'pPr')[0];
        if (pPr) {
            const pRPr = pPr.getElementsByTagNameNS(this.state.ns.w, 'rPr')[0];
            if (pRPr) {
                const sz = pRPr.getElementsByTagNameNS(this.state.ns.w, 'sz')[0];
                if (sz) { const val = sz.getAttribute('w:val'); if (val) return parseInt(val) / 2; }
            }
            const pStyle = pPr.getElementsByTagNameNS(this.state.ns.w, 'pStyle')[0];
            if (pStyle) {
                let currentStyleId: string | null | undefined = pStyle.getAttribute('w:val');
                while (currentStyleId) {
                    const style = this.state.docStyles.get(currentStyleId);
                    if (!style) break;
                    if (style.rPr && style.rPr.sz) {
                        const val = parseInt(style.rPr.sz);
                        if (!isNaN(val)) return val / 2;
                    }
                    currentStyleId = style.basedOn;
                }
            }
        }
        return 10;
    }

    private getListPrefix(numId: string, ilvl: string): string {
        const absRef = this.state.numberingMap.get(numId);
        if (!absRef) return '';
        const levels = this.state.abstractNumMap.get(absRef);
        if (!levels) return '';
        const lvl = levels.get(ilvl);
        if (!lvl) return '';
        // Counter tracking only; actual prefix rendered by HWPX numbering system
        const counterKey = `${numId}:${ilvl}`;
        if (this.state.listCounters) {
            if (this.state.listCounters.has(counterKey)) {
                this.state.listCounters.set(counterKey, this.state.listCounters.get(counterKey)! + 1);
            } else {
                this.state.listCounters.set(counterKey, lvl.start || 1);
            }
        }
        return '';
    }

    private getCellBorderFill(tcPr: Element | null, tblBorders: Element | null, r: number, c: number, rowCount: number, colCount: number, rowSpan = 1, colSpan = 1, styleTcBorders: Element | null = null, styleShd: Element | null = null): { id: string, [key: string]: any } {
        const borders: { [key: string]: BorderDef } = {
            left: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            right: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            top: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            bottom: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            slash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backSlash: { type: 'NONE', width: '0.1 mm', color: '#000000' }
        };

        const extract = (elem: Element | undefined): BorderDef | null => {
            if (!elem) return null;
            const val = elem.getAttribute('w:val');
            if (val === 'nil' || val === 'none') return { type: 'NONE', width: '0.1 mm', color: '#000000' };
            let type = 'SOLID';
            if (val === 'double') type = 'DOUBLE';
            else if (val === 'dashed') type = 'DASH';
            else if (val === 'dotted') type = 'DOT';
            const sz = parseInt(elem.getAttribute('w:sz') || '4');
            const mm = (sz / 8) * 0.3528;
            const presets = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0];
            let widthVal = presets[0];
            let minDiff = Math.abs(mm - presets[0]);
            for (let i = 1; i < presets.length; i++) {
                const diff = Math.abs(mm - presets[i]);
                if (diff < minDiff) { minDiff = diff; widthVal = presets[i]; }
            }
            if (mm > 5.0) widthVal = 5.0;
            const colorAttr = elem.getAttribute('w:color');
            let color = '#000000';
            if (colorAttr && colorAttr !== 'auto') color = `#${colorAttr}`;
            return { type, width: `${widthVal} mm`, color };
        };

        const mapBorder = (source: Element | null | undefined, side: string, targetSide: string) => {
            if (!source) return;
            const borderElem = source.getElementsByTagNameNS(this.state.ns.w, side)[0];
            if (borderElem) {
                const info = extract(borderElem);
                if (info) borders[targetSide] = info;
            }
        };

        // 1. 테이블 수준 테두리 (insideH/insideV → 내부 기본값, 위치별 outer 테두리)
        if (tblBorders) {
            const insideH = tblBorders.getElementsByTagNameNS(this.state.ns.w, 'insideH')[0];
            const insideV = tblBorders.getElementsByTagNameNS(this.state.ns.w, 'insideV')[0];
            const defH = extract(insideH);
            const defV = extract(insideV);
            if (defH) { borders.top = { ...defH }; borders.bottom = { ...defH }; }
            if (defV) { borders.left = { ...defV }; borders.right = { ...defV }; }
            if (r === 0) mapBorder(tblBorders, 'top', 'top');
            if (r + rowSpan >= rowCount) mapBorder(tblBorders, 'bottom', 'bottom');
            if (c === 0) mapBorder(tblBorders, 'left', 'left');
            if (c + colSpan >= colCount) mapBorder(tblBorders, 'right', 'right');
        }

        // 2. 셀 수준 테두리로 덮어쓰기
        let tcBorders = tcPr ? tcPr.getElementsByTagNameNS(this.state.ns.w, 'tcBorders')[0] : null;
        if (!tcBorders) tcBorders = styleTcBorders as any;
        if (tcBorders) {
            mapBorder(tcBorders, 'top', 'top');
            mapBorder(tcBorders, 'bottom', 'bottom');
            mapBorder(tcBorders, 'left', 'left');
            mapBorder(tcBorders, 'right', 'right');
            mapBorder(tcBorders, 'tl2br', 'slash');
            mapBorder(tcBorders, 'tr2bl', 'backSlash');
        }

        // 배경색
        let backColor = '#FFFFFF';
        let shd = tcPr ? tcPr.getElementsByTagNameNS(this.state.ns.w, 'shd')[0] : null;
        if (!shd) shd = styleShd as any;
        if (shd) {
            const fill = shd.getAttribute('w:fill');
            if (fill && fill !== 'auto') backColor = `#${fill}`;
        }

        const keyObj = {
            L: borders.left, R: borders.right, T: borders.top, B: borders.bottom,
            Sl: borders.slash, Bs: borders.backSlash, BG: backColor
        };
        const key = JSON.stringify(keyObj);

        for (const [, bf] of this.state.borderFills) {
            if (bf._key === key) return bf;
        }

        const newId = String(this.state.borderFillIdCounter++);
        const newBf = {
            id: newId,
            leftBorder: borders.left,
            rightBorder: borders.right,
            topBorder: borders.top,
            bottomBorder: borders.bottom,
            diagonal: null,
            slash: borders.slash,
            backSlash: borders.backSlash,
            backColor,
            _key: key
        };
        this.state.borderFills.set(newId, newBf);
        return newBf;
    }

    private convertTable(table: Element): string {
        const tableId = this.state.tableIdCounter++;

        const tblPr = table.getElementsByTagNameNS(this.state.ns.w, 'tblPr')[0];
        let tblBorders: Element | null = null;
        let tblStyleId: string | undefined;
        let defaultMargins = { left: 283, right: 283, top: 283, bottom: 283 };

        if (tblPr) {
            tblBorders = tblPr.getElementsByTagNameNS(this.state.ns.w, 'tblBorders')[0] || null;
            const styleNode = tblPr.getElementsByTagNameNS(this.state.ns.w, 'tblStyle')[0];
            if (styleNode) tblStyleId = styleNode.getAttribute('w:val') || undefined;
            const tblCellMar = tblPr.getElementsByTagNameNS(this.state.ns.w, 'tblCellMar')[0];
            if (tblCellMar) {
                const parseMar = (tag: string, def: number) => {
                    const m = tblCellMar.getElementsByTagNameNS(this.state.ns.w, tag)[0];
                    if (m) {
                        const w = parseInt(m.getAttribute('w:w') || '0');
                        return w * 5;
                    }
                    return def;
                };
                defaultMargins.left = parseMar('left', defaultMargins.left);
                defaultMargins.right = parseMar('right', defaultMargins.right);
                defaultMargins.top = parseMar('top', defaultMargins.top);
                defaultMargins.bottom = parseMar('bottom', defaultMargins.bottom);
            }
        }

        // DOCX <w:jc> → HWPX <hp:pos horzAlign="..."> 매핑
        // 단위 변환: DOCX dxa → HWPX HWPUnit : HWPUnit = dxa × 5
        let horzAlign = 'LEFT';
        let horzRelTo = 'COLUMN'; // 기본: 단(텍스트 영역) 기준
        if (tblPr) {
            const jcElem = tblPr.getElementsByTagNameNS(this.state.ns.w, 'jc')[0];
            if (jcElem) {
                const jcVal = jcElem.getAttribute('w:val') || 'left';
                if (jcVal === 'center') { horzAlign = 'CENTER'; horzRelTo = 'COLUMN'; }
                else if (jcVal === 'right') { horzAlign = 'RIGHT'; horzRelTo = 'COLUMN'; }
                else { horzAlign = 'LEFT'; horzRelTo = 'COLUMN'; }
            }
            // <w:tblpPr> 가 있으면 플로팅 표 → PAGE 기준 정렬 가능
            const tblpPr = tblPr.getElementsByTagNameNS(this.state.ns.w, 'tblpPr')[0];
            if (tblpPr) {
                const tblpAlign = tblpPr.getAttribute('w:tblpXSpec') || '';
                if (tblpAlign === 'center') { horzAlign = 'CENTER'; horzRelTo = 'PAGE'; }
                else if (tblpAlign === 'right') { horzAlign = 'RIGHT'; horzRelTo = 'PAGE'; }
            }
        }

        let styleTcBorders: Element | null = null;
        let styleShd: Element | null = null;

        if (tblStyleId) {
            let currentStyleId: string | undefined = tblStyleId;
            while (currentStyleId) {
                const style = this.state.docStyles.get(currentStyleId);
                if (!style) break;
                if (!tblBorders && style.tblBorders) tblBorders = style.tblBorders;
                if (!styleTcBorders && style.tcBorders) styleTcBorders = style.tcBorders;
                if (!styleShd && style.shd) styleShd = style.shd;
                currentStyleId = style.basedOn;
            }
        }

        const tblGrid = table.getElementsByTagNameNS(this.state.ns.w, 'tblGrid')[0];
        const gridCols = tblGrid ? Array.from(tblGrid.getElementsByTagNameNS(this.state.ns.w, 'gridCol')) : [];
        const colWidths: number[] = [];
        let totalWidth = 0;
        for (const col of gridCols) {
            const w = parseInt(col.getAttribute('w:w') || '0');
            const wHwp = w * 5;
            colWidths.push(wHwp);
            totalWidth += wHwp;
        }

        const rows = table.getElementsByTagNameNS(this.state.ns.w, 'tr');
        const rowCount = rows.length;
        const colCount = colWidths.length > 0 ? colWidths.length : 1;

        const cellMap: any[][] = [];
        for (let r = 0; r < rowCount; r++) {
            cellMap[r] = [];
            for (let c = 0; c < colCount; c++) {
                cellMap[r][c] = { occupied: false, rowSpan: 1, colSpan: 1, startRow: r, startCol: c, cell: null, vMerge: null };
            }
        }

        // 1. GridSpan & vMerge 파싱
        for (let r = 0; r < rowCount; r++) {
            const row = rows[r];
            const cells = row.getElementsByTagNameNS(this.state.ns.w, 'tc');
            let colIndex = 0;
            for (let i = 0; i < cells.length; i++) {
                const cell = cells[i];
                while (colIndex < colCount && cellMap[r][colIndex].occupied) colIndex++;
                if (colIndex >= colCount) break;
                const tcPr = cell.getElementsByTagNameNS(this.state.ns.w, 'tcPr')[0];
                let gridSpan = 1;
                let vMerge: string | null = null;
                if (tcPr) {
                    const gridSpanElem = tcPr.getElementsByTagNameNS(this.state.ns.w, 'gridSpan')[0];
                    if (gridSpanElem) gridSpan = parseInt(gridSpanElem.getAttribute('w:val') || '1');
                    const vMergeElem = tcPr.getElementsByTagNameNS(this.state.ns.w, 'vMerge')[0];
                    if (vMergeElem) vMerge = vMergeElem.getAttribute('w:val') || 'continue';
                }
                for (let k = 0; k < gridSpan; k++) {
                    if (colIndex + k < colCount) {
                        cellMap[r][colIndex + k].occupied = true;
                        cellMap[r][colIndex + k].cell = cell;
                        if (k === 0) {
                            cellMap[r][colIndex].colSpan = gridSpan;
                            cellMap[r][colIndex].vMerge = vMerge;
                        }
                    }
                }
                colIndex += gridSpan;
            }
        }

        // 2. vMerge 해결
        for (let c = 0; c < colCount; c++) {
            for (let r = 0; r < rowCount; r++) {
                const info = cellMap[r][c];
                if (info.cell && (info.vMerge === 'restart' || !info.vMerge)) {
                    let rowSpan = 1;
                    for (let nextR = r + 1; nextR < rowCount; nextR++) {
                        const nextInfo = cellMap[nextR][c];
                        if (nextInfo.cell && nextInfo.vMerge === 'continue') {
                            rowSpan++;
                            nextInfo.isMergedBelow = true;
                        } else break;
                    }
                    info.rowSpan = rowSpan;
                }
            }
        }

        if (totalWidth === 0) {
            totalWidth = this.state.TEXT_WIDTH;
            if (colWidths.length === 0 && colCount > 0) {
                const avgWidth = Math.floor(totalWidth / colCount);
                for (let i = 0; i < colCount; i++) colWidths.push(avgWidth);
            }
        }

        let tableHeight = 0;
        for (let r = 0; r < rowCount; r++) {
            const row = rows[r];
            let h = 0;
            const trPr = row.getElementsByTagNameNS(this.state.ns.w, 'trPr')[0];
            if (trPr) {
                const trHeightElem = trPr.getElementsByTagNameNS(this.state.ns.w, 'trHeight')[0];
                if (trHeightElem) h = parseInt(trHeightElem.getAttribute('w:val') || '0') * 5;
            }
            tableHeight += (h > 0 ? h : 1500);
        }

        let xml = `<hp:p id="${this.state.idCounter++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:tbl id="${tableId}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="${rowCount}" colCnt="${colCount}" cellSpacing="0" borderFillIDRef="3" noAdjust="0">`;
        xml += `<hp:sz width="${totalWidth}" widthRelTo="ABSOLUTE" height="${tableHeight}" heightRelTo="ABSOLUTE" protect="0"/>`;
        xml += `<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="${horzRelTo}" vertAlign="TOP" horzAlign="${horzAlign}" vertOffset="0" horzOffset="0"/>`;
        xml += `<hp:outMargin left="0" right="0" top="0" bottom="0"/>`;
        xml += `<hp:inMargin left="540" right="540" top="0" bottom="0"/>`;

        for (let r = 0; r < rowCount; r++) {
            xml += `<hp:tr>`;

            let c = 0;
            while (c < colCount) {
                const info = cellMap[r][c];
                if (info.isMergedBelow) { c++; continue; }

                const cell = info.cell as Element;
                let cellWidth = 0;
                for (let k = 0; k < info.colSpan; k++) cellWidth += colWidths[c + k] || 0;

                if (!cell) {
                    xml += `<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="4">`;
                    xml += `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">`;
                    xml += `<hp:p id="${this.state.idCounter++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t></hp:t></hp:run><hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="0" horzpos="0" horzsize="${cellWidth}" flags="393216"/></hp:linesegarray></hp:p>`;
                    xml += `</hp:subList>`;
                    xml += `<hp:cellAddr colAddr="${c}" rowAddr="${r}"/>`;
                    xml += `<hp:cellSpan colSpan="${info.colSpan}" rowSpan="${info.rowSpan}"/>`;
                    xml += `<hp:cellSz width="${cellWidth}" height="0"/>`;
                    xml += `<hp:cellMargin left="255" right="255" top="141" bottom="141"/>`;
                    xml += `</hp:tc>`;
                    c++;
                    continue;
                }

                const tcPr = cell.getElementsByTagNameNS(this.state.ns.w, 'tcPr')[0];
                // JS 원본과 동일하게 tblBorders와 셀 위치 정보를 전달
                const borderFill = this.getCellBorderFill(tcPr || null, tblBorders, r, c, rowCount, colCount, info.rowSpan, info.colSpan, styleTcBorders, styleShd);

                // 셀 마진 (테이블 기본값 → 셀 직접 속성으로 재정의)
                let margins = { ...defaultMargins };
                if (tcPr) {
                    const tcMar = tcPr.getElementsByTagNameNS(this.state.ns.w, 'tcMar')[0];
                    if (tcMar) {
                        const parseMar = (tag: string) => {
                            const mTag = tcMar.getElementsByTagNameNS(this.state.ns.w, tag)[0];
                            if (mTag) return parseInt(mTag.getAttribute('w:w') || '0') * 5;
                            return null;
                        };
                        const l = parseMar('left'); if (l !== null) margins.left = l;
                        const rv = parseMar('right'); if (rv !== null) margins.right = rv;
                        const t = parseMar('top'); if (t !== null) margins.top = t;
                        const b = parseMar('bottom'); if (b !== null) margins.bottom = b;
                    }
                }

                const hasTcMar = (tcPr && tcPr.getElementsByTagNameNS(this.state.ns.w, 'tcMar')[0]) ? '1' : '0';

                let vertAlign = 'TOP';
                if (tcPr) {
                    const vAlign = tcPr.getElementsByTagNameNS(this.state.ns.w, 'vAlign')[0];
                    if (vAlign) {
                        const val = vAlign.getAttribute('w:val');
                        if (val === 'center') vertAlign = 'CENTER';
                        else if (val === 'bottom') vertAlign = 'BOTTOM';
                    }
                }

                xml += `<hp:tc name="" header="0" hasMargin="${hasTcMar}" protect="0" editable="0" dirty="0" borderFillIDRef="${borderFill.id}">`;
                xml += `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${vertAlign}" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">`;

                let hasContent = false;
                for (let i = 0; i < cell.childNodes.length; i++) {
                    const childNode = cell.childNodes[i] as Element;
                    if (childNode.nodeName === 'w:p') {
                        xml += this.convertParagraph(childNode);
                        hasContent = true;
                    } else if (childNode.nodeName === 'w:tbl') {
                        xml += this.convertTable(childNode);
                        hasContent = true;
                    }
                }
                // 빈 셀 안전장치
                if (!hasContent) {
                    xml += `<hp:p id="${this.state.idCounter++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t></hp:t></hp:run><hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="0" horzpos="0" horzsize="${cellWidth}" flags="393216"/></hp:linesegarray></hp:p>`;
                }

                xml += `</hp:subList>`;
                xml += `<hp:cellAddr colAddr="${c}" rowAddr="${r}"/>`;
                xml += `<hp:cellSpan colSpan="${info.colSpan}" rowSpan="${info.rowSpan}"/>`;
                xml += `<hp:cellSz width="${cellWidth}" height="0"/>`;
                xml += `<hp:cellMargin left="${margins.left}" right="${margins.right}" top="${margins.top}" bottom="${margins.bottom}"/>`;
                xml += `</hp:tc>`;

                c += info.colSpan;
            }
            xml += `</hp:tr>`;
        }

        xml += `</hp:tbl></hp:run>`;
        // JS 원본과 동일하게 표 문단에 linesegarray 추가 및 currentVertPos 업데이트
        xml += `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${this.state.currentVertPos}" vertsize="${tableHeight}" textheight="${tableHeight}" baseline="${tableHeight}" spacing="0" horzpos="0" horzsize="${totalWidth}" flags="393216"/></hp:linesegarray>`;
        xml += `</hp:p>`;
        this.state.currentVertPos += tableHeight;

        return xml;
    }

    private convertDrawing(drawing: Element): string {
        // 1. Pictures (blip)
        // 수정된 부분: 네임스페이스의 방해를 받지 않도록 모든 요소 중 localName이 'blip'인 것을 필터링
        const allElements = drawing.getElementsByTagName('*');
        const blips = Array.from(allElements).filter(el => el.localName === 'blip');

        // 결과 확인용 로그 (확인 후 삭제하셔도 됩니다)
        console.log(`✅ 찾은 이미지(blip) 개수: ${blips.length}`);
        if (blips.length > 0) {
            const embedId = blips[0].getAttributeNS(this.state.ns.r, 'embed') || blips[0].getAttribute('r:embed') || blips[0].getAttribute('embed');
            console.log(`- 추출된 이미지 연결 ID: ${embedId}`);

            return this.convertPicture(drawing, blips[0]);
        }

        // 2. Shapes (wps:wsp)
        // DOCX shapes are complex. We will approximate them as rectangles with checking allow overlap etc.
        const textboxes = drawing.getElementsByTagNameNS('http://schemas.microsoft.com/office/word/2010/wordprocessingShape', 'txbx');
        let textContent = '';
        if (textboxes.length > 0) {
            const paras = textboxes[0].getElementsByTagNameNS(this.state.ns.w, 'p');
            for (const p of paras) {
                const tNodes = p.getElementsByTagNameNS(this.state.ns.w, 't');
                for (const t of tNodes) textContent += t.textContent + '\n';
            }
        }

        if (textContent) {
            return `<hp:run charPrIDRef="0"><hp:t>[텍스트 상자: ${escapeXml(textContent.trim())}]</hp:t></hp:run>`;
        }

        // 3. Charts (c:chart)
        const charts = drawing.getElementsByTagName('c:chart');
        if (charts.length > 0 || drawing.textContent?.includes('chart')) {
            return `<hp:run charPrIDRef="0"><hp:t>[차트: 변환되지 않음]</hp:t></hp:run>`;
        }

        return '';
    }

    private convertPicture(drawing: Element, blip: Element): string {
        // JS 원본과 동일: getAttributeNS로 r:embed 접근, 폴백 추가
        const embed = blip.getAttributeNS(this.state.ns.r, 'embed') || blip.getAttribute('r:embed') || blip.getAttribute('embed');
        if (!embed) return '';
        const imgInfo = this.state.images.get(embed);
        console.log(this.state.images);
        if (!imgInfo) return '';

        const picId = this.state.tableIdCounter++;

        // 이미지 크기 추출 (EMU → HWPUNIT)
        let width = 27577; // 기본값
        let height = 19913;

        const extents = drawing.getElementsByTagNameNS(this.state.ns.wp, 'extent');
        if (extents.length > 0) {
            const cxAttr = extents[0].getAttribute('cx');
            const cyAttr = extents[0].getAttribute('cy');
            if (cxAttr && cyAttr) {
                width = Math.round(parseInt(cxAttr) * this.state.HWPUNIT_PER_INCH / this.state.EMU_PER_INCH);
                height = Math.round(parseInt(cyAttr) * this.state.HWPUNIT_PER_INCH / this.state.EMU_PER_INCH);
            }
        }

        const instId = Math.floor(Math.random() * 2000000000);

        // 인라인 vs 플로팅 판별
        const isInline = drawing.getElementsByTagNameNS(this.state.ns.wp, 'inline').length > 0;
        const anchor = drawing.getElementsByTagNameNS(this.state.ns.wp, 'anchor')[0];

        // 기본 위치 속성
        let treatAsChar = isInline ? '1' : '0';
        let textWrap = 'TOP_AND_BOTTOM';
        let vertRelTo = 'PARA';
        let horzRelTo = isInline ? 'COLUMN' : 'COLUMN';
        let vertAlign = 'TOP';
        let horzAlign = 'LEFT';
        let vertOffset = '0';
        let horzOffset = '0';
        let outMarginLeft = '0';
        let outMarginRight = '0';
        let outMarginTop = '0';
        let outMarginBottom = '0';

        // 플로팅 이미지 속성 파싱
        if (anchor) {
            // textWrap 결정
            const wrapSquare = anchor.getElementsByTagNameNS(this.state.ns.wp, 'wrapSquare')[0];
            const wrapTight = anchor.getElementsByTagNameNS(this.state.ns.wp, 'wrapTight')[0];
            const wrapThrough = anchor.getElementsByTagNameNS(this.state.ns.wp, 'wrapThrough')[0];
            const wrapNone = anchor.getElementsByTagNameNS(this.state.ns.wp, 'wrapNone')[0];
            const wrapTopBottom = anchor.getElementsByTagNameNS(this.state.ns.wp, 'wrapTopAndBottom')[0];

            if (wrapSquare) textWrap = 'SQUARE';
            else if (wrapTight || wrapThrough) textWrap = 'TIGHT';
            else if (wrapNone) textWrap = 'BEHIND_TEXT';
            else if (wrapTopBottom) textWrap = 'TOP_AND_BOTTOM';

            // 수직 위치 파싱
            const posV = anchor.getElementsByTagNameNS(this.state.ns.wp, 'positionV')[0];
            if (posV) {
                const relFrom = posV.getAttribute('relativeFrom');
                if (relFrom === 'page') vertRelTo = 'PAGE';
                else if (relFrom === 'paragraph') vertRelTo = 'PARA';
                else if (relFrom === 'margin') vertRelTo = 'PAPER';

                const posOffset = posV.getElementsByTagNameNS(this.state.ns.wp, 'posOffset')[0];
                if (posOffset) {
                    // EMU → HWPUNIT
                    vertOffset = String(Math.round(parseInt(posOffset.textContent || '0') * this.state.HWPUNIT_PER_INCH / this.state.EMU_PER_INCH));
                }
                const align = posV.getElementsByTagNameNS(this.state.ns.wp, 'align')[0];
                if (align) {
                    const val = align.textContent;
                    if (val === 'center') vertAlign = 'CENTER';
                    else if (val === 'bottom') vertAlign = 'BOTTOM';
                    else if (val === 'top') vertAlign = 'TOP';
                }
            }

            // 수평 위치 파싱
            const posH = anchor.getElementsByTagNameNS(this.state.ns.wp, 'positionH')[0];
            if (posH) {
                const relFrom = posH.getAttribute('relativeFrom');
                if (relFrom === 'page') horzRelTo = 'PAGE';
                else if (relFrom === 'column') horzRelTo = 'COLUMN';
                else if (relFrom === 'margin') horzRelTo = 'PAPER';

                const posOffset = posH.getElementsByTagNameNS(this.state.ns.wp, 'posOffset')[0];
                if (posOffset) {
                    horzOffset = String(Math.round(parseInt(posOffset.textContent || '0') * this.state.HWPUNIT_PER_INCH / this.state.EMU_PER_INCH));
                }
                const align = posH.getElementsByTagNameNS(this.state.ns.wp, 'align')[0];
                if (align) {
                    const val = align.textContent;
                    if (val === 'center') horzAlign = 'CENTER';
                    else if (val === 'right') horzAlign = 'RIGHT';
                    else if (val === 'left') horzAlign = 'LEFT';
                }
            }

            // 외부 여백 (effectExtent)
            const effectExtent = anchor.getElementsByTagNameNS(this.state.ns.wp, 'effectExtent')[0];
            if (effectExtent) {
                const emuToHwp = (val: string | null) => String(Math.round(parseInt(val || '0') * this.state.HWPUNIT_PER_INCH / this.state.EMU_PER_INCH));
                outMarginLeft = emuToHwp(effectExtent.getAttribute('l'));
                outMarginRight = emuToHwp(effectExtent.getAttribute('r'));
                outMarginTop = emuToHwp(effectExtent.getAttribute('t'));
                outMarginBottom = emuToHwp(effectExtent.getAttribute('b'));
            }
        }

        // HWPX 이미지 XML 생성
        let xml = `<hp:run charPrIDRef="0">`;
        xml += `<hp:pic id="${picId}" zOrder="0" numberingType="NONE" textWrap="${textWrap}" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instId="${instId}" reverse="0">`; xml += `<hp:offset x="0" y="0"/>`;
        xml += `<hp:orgSz width="${width}" height="${height}"/>`;
        xml += `<hp:curSz width="0" height="0"/>`;
        xml += `<hp:flip horizontal="0" vertical="0"/>`;
        xml += `<hp:rotationInfo angle="0" centerX="${Math.round(width / 2)}" centerY="${Math.round(height / 2)}" rotateimage="1"/>`;
        xml += `<hp:renderingInfo>`;
        xml += `<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>`;
        xml += `<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>`;
        xml += `<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>`;
        xml += `</hp:renderingInfo>`;
        xml += `<hp:imgRect>`;
        xml += `<hc:pt0 x="0" y="0"/>`;
        xml += `<hc:pt1 x="${width}" y="0"/>`;
        xml += `<hc:pt2 x="${width}" y="${height}"/>`;
        xml += `<hc:pt3 x="0" y="${height}"/>`;
        xml += `</hp:imgRect>`;
        xml += `<hp:imgClip left="0" right="0" top="0" bottom="0"/>`; // 참조: 항상 0
        xml += `<hp:inMargin left="0" right="0" top="0" bottom="0"/>`;
        xml += `<hc:img binaryItemIDRef="${imgInfo.manifestId}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>`;
        xml += `<hp:effects/>`;
        xml += `<hp:sz width="${width}" widthRelTo="ABSOLUTE" height="${height}" heightRelTo="ABSOLUTE" protect="0"/>`;
        xml += `<hp:pos treatAsChar="${treatAsChar}" affectLSpacing="0" flowWithText="1" allowOverlap="1" holdAnchorAndSO="0" vertRelTo="${vertRelTo}" horzRelTo="${horzRelTo}" vertAlign="${vertAlign}" horzAlign="${horzAlign}" vertOffset="${vertOffset}" horzOffset="${horzOffset}"/>`;
        xml += `<hp:outMargin left="${outMarginLeft}" right="${outMarginRight}" top="${outMarginTop}" bottom="${outMarginBottom}"/>`;
        xml += `<hp:shapeComment>그림입니다.\n원본 그림의 이름: ${imgInfo.manifestId}.${imgInfo.ext}\n원본 그림의 크기: 가로 ${Math.round(width / 283.465)}mm, 세로 ${Math.round(height / 283.465)}mm</hp:shapeComment>`;
        xml += `</hp:pic></hp:run>`;

        return xml;
    }



    private createParaProperties() {
        let xml = `<hh:paraProperties itemCnt="${this.state.paraShapes.size}">`;
        for (const [id, ps] of this.state.paraShapes) {
            xml += `<hh:paraPr id="${id}" tabPrIDRef="${ps.tabPrIDRef || '1'}" condense="0" fontLineHeight="0" snapToGrid="1" suppressLineNumbers="0" checked="0">`;
            xml += `<hh:align horizontal="${ps.align}" vertical="BASELINE"/>`;
            xml += `<hh:heading type="${ps.heading}" idRef="${ps.headingIdRef || '0'}" level="${ps.level || '0'}"/>`;
            // Simplified breakSetting
            xml += `<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" widowOrphan="0" keepWithNext="${ps.keepWithNext || '0'}" keepLines="${ps.keepLines || '0'}" pageBreakBefore="${ps.pageBreakBefore || '0'}" lineWrap="BREAK"/>`;
            xml += `<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>`;
            xml += `<hh:margin><hc:intent value="${ps.indent || 0}" unit="HWPUNIT"/><hc:left value="${ps.leftMargin || 0}" unit="HWPUNIT"/><hc:right value="${ps.rightMargin || 0}" unit="HWPUNIT"/><hc:prev value="${ps.prevSpacing || 0}" unit="HWPUNIT"/><hc:next value="${ps.nextSpacing || 0}" unit="HWPUNIT"/></hh:margin>`;
            xml += `<hh:lineSpacing type="${ps.lineSpacingType}" value="${ps.lineSpacingVal}" unit="HWPUNIT"/>`;
            xml += `<hh:border borderFillIDRef="${ps.borderFillIDRef}" offsetLeft="400" offsetRight="400" offsetTop="100" offsetBottom="100" connect="0" ignoreMargin="0"/>`;
            xml += `</hh:paraPr>`;
        }
        xml += `</hh:paraProperties>`;
        return xml;
    }

    private createCharProperties() {
        let xml = `<hh:charProperties itemCnt="${this.state.charShapes.size}">`;
        for (const [id, cs] of this.state.charShapes) {
            xml += `<hh:charPr id="${id}" height="${cs.height}" textColor="${cs.textColor}" shadeColor="${cs.shadeColor}" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="${cs.borderFillIDRef}">`;
            const fIds = cs.fontIds || { hangulId: 0, latinId: 0, hanjaId: 0, japaneseId: 0, otherId: 0, symbolId: 0, userId: 0 };
            xml += `<hh:fontRef hangul="${fIds.hangulId}" latin="${fIds.latinId}" hanja="${fIds.hanjaId}" japanese="${fIds.japaneseId}" other="${fIds.otherId}" symbol="${fIds.symbolId}" user="${fIds.userId}"/>`;
            xml += `<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>`;
            xml += `<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>`;
            xml += `<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>`;
            xml += `<hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>`;
            if (cs.bold) xml += `<hh:bold/>`;
            if (cs.italic) xml += `<hh:italic/>`;
            if (cs.underline) xml += `<hh:underline type="BOTTOM" shape="SOLID" color="${cs.textColor}"/>`;
            if (cs.strike) xml += `<hh:strikeout shape="SOLID" color="${cs.textColor}"/>`;
            if (cs.supscript) xml += `<hh:supscript/>`;
            if (cs.subscript) xml += `<hh:subscript/>`;
            xml += `</hh:charPr>`;
        }
        xml += `</hh:charProperties>`;
        return xml;
    }

    private createContentHpf() {
        let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n`;
        xml += `<opf:package xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" version="" unique-identifier="" id="">\n`;

        // 메타데이터 (기존 코드 유지)
        xml += `<opf:metadata>\n`;
        xml += `  <opf:title>문서</opf:title>\n`;
        xml += `  <opf:language>ko</opf:language>\n`;
        xml += `  <opf:meta name="subject" content="text"></opf:meta>\n`;
        xml += `</opf:metadata>\n`;

        xml += `<opf:manifest>\n`;

        for (const [, imgInfo] of this.state.images) {
            const mediaType = imgInfo.ext === 'png' ? 'image/png' :
                (imgInfo.ext === 'jpg' || imgInfo.ext === 'jpeg' ? 'image/jpeg' :
                    (imgInfo.ext === 'gif' ? 'image/gif' : 'image/png'));
            // href 부분에 ../ 가 들어가는 것이 핵심입니다!
            xml += `  <opf:item id="${imgInfo.manifestId}" href="BinData/${imgInfo.manifestId}.${imgInfo.ext}" media-type="${mediaType}" isEmbeded="1"/>\n`;
        }

        xml += `  <opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>\n`;
        xml += `  <opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>\n`;
        xml += `  <opf:item id="settings" href="settings.xml" media-type="application/xml"/>\n`;

        xml += `</opf:manifest>\n`;

        xml += `<opf:spine>\n`;
        xml += `  <opf:itemref idref="header"/>\n`;
        xml += `  <opf:itemref idref="section0"/>\n`;
        xml += `</opf:spine>\n`;

        xml += `</opf:package>`;
        return xml;
    }


    private createContainer() {
        // 참조 형식에 맞춘 단일 줄 compact XML
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/><ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/><ocf:rootfile full-path="META-INF/container.rdf" media-type="application/rdf+xml"/></ocf:rootfiles></ocf:container>`;
    }

    private createContainerhdf() {
        // container.rdf - 참조 형식에 맞춘 compact XML
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/header.xml"/></rdf:Description><rdf:Description rdf:about="Contents/header.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/></rdf:Description><rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/section0.xml"/></rdf:Description><rdf:Description rdf:about="Contents/section0.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/></rdf:Description><rdf:Description rdf:about=""><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"/></rdf:Description></rdf:RDF>`;
    }

    private createManifest() {
        let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n`;
        // HWPX 표준에 맞게 manifest 네임스페이스 선언
        xml += `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">\n`;
        xml += `  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/hwpml-package+xml"/>\n`;
        xml += `  <manifest:file-entry manifest:full-path="version.xml" manifest:media-type="application/xml"/>\n`;
        xml += `  <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="application/xml"/>\n`;
        xml += `  <manifest:file-entry manifest:full-path="Contents/content.hpf" manifest:media-type="application/hwpml-package+xml"/>\n`;
        xml += `  <manifest:file-entry manifest:full-path="Contents/header.xml" manifest:media-type="application/xml"/>\n`;
        xml += `  <manifest:file-entry manifest:full-path="Contents/section0.xml" manifest:media-type="application/xml"/>\n`;

        // 🔥 여기가 핵심! BinData 안의 이미지들을 명부에 차례대로 등록합니다.
        for (const [, imgInfo] of this.state.images) {
            const mediaType = imgInfo.ext === 'png' ? 'image/png' :
                (imgInfo.ext === 'jpg' || imgInfo.ext === 'jpeg' ? 'image/jpeg' :
                    (imgInfo.ext === 'gif' ? 'image/gif' : 'image/png'));
            xml += `  <manifest:file-entry manifest:full-path="Contents/${imgInfo.path}" manifest:media-type="${mediaType}"/>\n`;
        }

        xml += `</manifest:manifest>`;
        return xml;
    }

    public createHwpx(doc: Document): any {
        // 1. Convert Section first to populate style maps (charShapes, borderFills, etc.)
        const section0 = this.createSection(doc);

        // 2. Create Header (now it includes all styles found during conversion)
        const header = this.createHeader();

        return {
            mimetype: 'application/hwp+zip',
            version: this.createVersion(),
            settings: this.createSettings(),
            header: header,
            section0: section0,
            contentHpf: this.createContentHpf(),
            container: this.createContainer(),
            containerRdf: this.createContainerhdf(),
            manifest: this.createManifest()
        };
    }

    private createVersion() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hv:HCFVersion xmlns:hv="${this.state.hwpxNs.hv}" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="0" buildNumber="1" os="6" xmlVersion="1.2" application="Hancom Office Hangul" appVersion="11, 20, 0, 1520"/>`;
    }

    private createSettings() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ha:HWPApplicationSetting xmlns:ha="${this.state.hwpxNs.ha}" xmlns:config="${this.state.hwpxNs.config}"><ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/></ha:HWPApplicationSetting>`;
    }

    private createHeader() {
        let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hh:head xmlns:ha="${this.state.hwpxNs.ha}" xmlns:hp="${this.state.hwpxNs.hp}" xmlns:hp10="${this.state.hwpxNs.hp10}" xmlns:hs="${this.state.hwpxNs.hs}" xmlns:hc="${this.state.hwpxNs.hc}" xmlns:hh="${this.state.hwpxNs.hh}" xmlns:hhs="${this.state.hwpxNs.hhs}" xmlns:hm="${this.state.hwpxNs.hm}" xmlns:hpf="${this.state.hwpxNs.hpf}" xmlns:dc="${this.state.hwpxNs.dc}" xmlns:opf="${this.state.hwpxNs.opf}" xmlns:ooxmlchart="${this.state.hwpxNs.ooxmlchart}" xmlns:epub="${this.state.hwpxNs.epub}" xmlns:config="${this.state.hwpxNs.config}" version="1.2" secCnt="1">`;
        xml += `<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>`;
        xml += `<hh:refList>`;
        xml += this.createFontFaces();
        xml += this.createBorderFills();
        xml += this.createBinDataProperties();
        xml += this.createCharProperties();
        xml += this.createTabProperties();
        xml += this.createNumberings();
        xml += this.createParaProperties();
        xml += this.createStyles();
        xml += `<hh:memoProperties itemCnt="1"><hh:memoPr id="1" width="15024" lineWidth="1" lineType="SOLID" lineColor="#000000" fillColor="#CCFF99" activeColor="#FFFF99" memoType="NOMAL"/></hh:memoProperties>`;
        xml += `</hh:refList>`;
        xml += `    <hh:compatibleDocument targetProgram="MS_WORD">
        <hh:layoutCompatibility>
            <hh:applyFontWeightToBold />
            <hh:useInnerUnderline />
            <hh:useLowercaseStrikeout />
            <hh:extendLineheightToOffset />
            <hh:treatQuotationAsLatin />
            <hh:doNotAlignWhitespaceOnRight />
            <hh:doNotAdjustWordInJustify />
            <hh:baseCharUnitOnEAsian />
            <hh:baseCharUnitOfIndentOnFirstChar />
            <hh:adjustLineheightToFont />
            <hh:adjustBaselineInFixedLinespacing />
            <hh:applyPrevspacingBeneathObject />
            <hh:applyNextspacingOfLastPara />
            <hh:adjustParaBorderfillToSpacing />
            <hh:connectParaBorderfillOfEqualBorder />
            <hh:adjustParaBorderOffsetWithBorder />
            <hh:extendLineheightToParaBorderOffset />
            <hh:applyParaBorderToOutside />
            <hh:applyMinColumnWidthTo1mm />
            <hh:applyTabPosBasedOnSegment />
            <hh:breakTabOverLine />
            <hh:adjustVertPosOfLine />
            <hh:doNotAlignLastForbidden />
            <hh:adjustMarginFromAdjustLineheight />
            <hh:baseLineSpacingOnLineGrid />
            <hh:applyCharSpacingToCharGrid />
            <hh:doNotApplyGridInHeaderFooter />
            <hh:applyExtendHeaderFooterEachSection />
            <hh:doNotApplyLinegridAtNoLinespacing />
            <hh:doNotAdjustEmptyAnchorLine />
            <hh:overlapBothAllowOverlap />
            <hh:extendVertLimitToPageMargins />
            <hh:doNotHoldAnchorOfTable />
            <hh:doNotFormattingAtBeneathAnchor />
            <hh:adjustBaselineOfObjectToBottom />
        </hh:layoutCompatibility>
    </hh:compatibleDocument>`;
        xml += `<hh:docOption><hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/></hh:docOption>`;
        xml += `<hh:trackchageConfig flags="0"/>`;
        xml += `</hh:head>`;
        return xml;
    }

    private createBorderFills() {
        let xml = `<hh:borderFills itemCnt="${this.state.borderFills.size}">`;
        for (const [id, bf] of this.state.borderFills) {
            const backColor = bf.backColor || '#FFFFFF'; // 기본 흰색

            // 대각선
            const slash = bf.slash || { type: 'NONE', width: '0.1 mm', color: '#000000' };
            const backSlash = bf.backSlash || { type: 'NONE', width: '0.1 mm', color: '#000000' };

            xml += `<hh:borderFill id="${id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">`;
            xml += `<hh:slash type="${slash.type}" Crooked="0" isCounter="0" />`;
            xml += `<hh:backSlash type="${backSlash.type}" Crooked="0" isCounter="0" />`;
            xml += `<hh:leftBorder type="${bf.leftBorder.type}" width="${bf.leftBorder.width}" color="${bf.leftBorder.color}" />`;
            xml += `<hh:rightBorder type="${bf.rightBorder.type}" width="${bf.rightBorder.width}" color="${bf.rightBorder.color}" />`;
            xml += `<hh:topBorder type="${bf.topBorder.type}" width="${bf.topBorder.width}" color="${bf.topBorder.color}" />`;
            xml += `<hh:bottomBorder type="${bf.bottomBorder.type}" width="${bf.bottomBorder.width}" color="${bf.bottomBorder.color}" />`;
            xml += `<hh:diagonal type="SOLID" width="0.1 mm" color="#000000" />`;
            // fillBrush: input 원본과 동일한 hc:fillBrush > hc:winBrush 구조
            xml += `<hc:fillBrush><hc:winBrush faceColor="${backColor}" hatchColor="#000000" alpha="0" /></hc:fillBrush>`;
            xml += `</hh:borderFill>`;
        }
        xml += `</hh:borderFills>`;
        return xml;
    }

    private createBinDataProperties() {
        if (this.state.images.size === 0) return '';
        let xml = `<hh:binDataProperties itemCnt="${this.state.images.size}">`;
        for (const [, imgInfo] of this.state.images) {
            // id는 본문에서 참조할 번호, itemIDRef는 HWPX manifest의 ID입니다.
            xml += `<hh:binData id="${imgInfo.binDataId}" itemIDRef="${imgInfo.manifestId}" extension="${imgInfo.ext}" format="${imgInfo.ext}" type="EMBEDDING" state="NOT_ACCESSED"/>`;
        }
        xml += `</hh:binDataProperties>`;
        return xml;
    }

    private createTabProperties() {
        // PDF 표 36 참고 - 2개 탭 속성 정의 (기본 + 헤더/푸터용)
        let xml = `<hh:tabProperties itemCnt="2">`;
        xml += `<hh:tabPr id="0" itemCnt="0" autoTabLeft="1" autoTabRight="0"/>`;
        // 탭 속성 1: 중앙/오른쪽 정렬 탭 (헤더/푸터용)
        xml += `<hh:tabPr id="1" itemCnt="2" autoTabLeft="0" autoTabRight="0">`;
        const centerTab = Math.round((this.state.PAGE_WIDTH - this.state.MARGIN_LEFT - this.state.MARGIN_RIGHT) / 2);
        const rightTab = this.state.PAGE_WIDTH - this.state.MARGIN_LEFT - this.state.MARGIN_RIGHT;
        xml += `<hh:tabItem pos="${centerTab}" type="CENTER" leader="NONE"/>`;
        xml += `<hh:tabItem pos="${rightTab}" type="RIGHT" leader="NONE"/>`;
        xml += `</hh:tabPr>`;
        xml += `</hh:tabProperties>`;
        return xml;
    }

    private createNumberings() {
        // 넘버링 정의가 없으면 빈 요소 반환
        if (this.state.abstractNumMap.size === 0) {
            return `<hh:numberings itemCnt="0"/>`;
        }

        // DOCX numFmt → HWPX numFormat 변환 맵
        const numFmtMap: { [key: string]: string } = {
            'decimal': 'DIGIT',
            'lowerRoman': 'ROMAN_SMALL',
            'upperRoman': 'ROMAN_CAPITAL',
            'lowerLetter': 'LATIN_SMALL',
            'upperLetter': 'LATIN_CAPITAL',
            'bullet': 'BULLET',
            'ganada': 'DIGIT',
            'chosung': 'DIGIT',
            'none': 'DIGIT'
        };

        // bullet 기호 매핑
        const bulletSymbolMap: { [key: string]: string } = {
            '': '·',
            'o': 'o',
            '§': '§',
        };

        let xml = `<hh:numberings itemCnt="${this.state.abstractNumMap.size}">`;
        let numberingId = 1;

        for (const [_absId, levels] of this.state.abstractNumMap) {
            xml += `<hh:numbering id="${numberingId}" start="${levels.get('0')?.start || 1}">`;

            // 10개 레벨 생성 (HWPX 기본 구조)
            for (let level = 1; level <= 10; level++) {
                const lvlData = levels.get(String(level - 1));

                if (lvlData) {
                    const numFormat = numFmtMap[lvlData.numFmt] || 'DIGIT';
                    const startVal = lvlData.start || 1;

                    // 들여쓰기 계산 (DOCX twip → HWPUNIT, 1twip = 2HWPUNIT)
                    // const indent = lvlData.lvlLeft ? Math.round(parseInt(lvlData.lvlLeft) * 2) : (level * 2200);
                    // const hanging = lvlData.lvlHanging ? Math.round(parseInt(lvlData.lvlHanging) * 2) : 3600;

                    // bullet인 경우
                    if (lvlData.numFmt === 'bullet') {
                        const symbol = bulletSymbolMap[lvlData.lvlText] || '·';
                        xml += `<hh:paraHead start="${startVal}" level="${level}" align="LEFT" useInstWidth="1" autoIndent="1"`;
                        xml += ` widthAdjust="0" textOffsetType="PERCENT" textOffset="100" numFormat="BULLET"`;
                        xml += ` charPrIDRef="0" checkable="0">${symbol}</hh:paraHead>`;
                    } else {
                        // 숫자/문자 번호 매기기
                        // lvlText에서 HWPX 포맷 문자 생성
                        let headText = this.convertLvlTextToHwpx(lvlData.lvlText, numFormat, level);

                        xml += `<hh:paraHead start="${startVal}" level="${level}" align="LEFT" useInstWidth="1" autoIndent="1"`;
                        xml += ` widthAdjust="0" textOffsetType="PERCENT" textOffset="100" numFormat="${numFormat}"`;
                        xml += ` charPrIDRef="0" checkable="0">${escapeXml(headText)}</hh:paraHead>`;
                    }
                } else {
                    // 정의되지 않은 레벨은 기본 DIGIT 사용
                    xml += `<hh:paraHead start="1" level="${level}" align="LEFT" useInstWidth="1" autoIndent="0"`;
                    xml += ` widthAdjust="0" textOffsetType="PERCENT" textOffset="100" numFormat="DIGIT"`;
                    xml += ` charPrIDRef="0" checkable="0">^N</hh:paraHead>`;
                }
            }

            xml += `</hh:numbering>`;
            numberingId++;
        }

        xml += `</hh:numberings>`;
        return xml;
    }

    private convertLvlTextToHwpx(lvlText: string, _numFormat: string, level: number) {
        if (!lvlText) return '^N.';

        // %1, %2 등의 플레이스홀더를 HWPX ^N 포맷으로 변환
        let result = lvlText;

        // 현재 레벨의 플레이스홀더(%level)를 ^N으로 치환
        const currentPlaceholder = `%${level}`;
        result = result.replace(currentPlaceholder, '^N');

        // 다른 레벨의 플레이스홀더는 제거 (HWPX는 다단계 참조를 다르게 처리)
        for (let i = 1; i <= 10; i++) {
            if (i !== level) {
                result = result.replace(`%${i}`, '');
            }
        }

        return result;
    }

    private createStyles() {
        // DOCX 스타일을 HWPX 스타일로 의미있게 매핑
        // MS Word 호환 스타일 62개 생성 (input 원본과 동일한 구조)
        const styleDefinitions = [
            { id: 0, type: 'PARA', name: '바탕글', engName: 'Normal', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 0 },
            { id: 1, type: 'CHAR', name: 'Book Title', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 1 },
            { id: 2, type: 'PARA', name: 'Caption', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 2 },
            { id: 3, type: 'CHAR', name: 'Default Paragraph Font', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 3 },
            { id: 4, type: 'CHAR', name: 'Emphasis', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 4 },
            { id: 5, type: 'CHAR', name: 'Endnote Text Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 5 },
            { id: 6, type: 'CHAR', name: 'FollowedHyperlink', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 6 },
            { id: 7, type: 'PARA', name: 'Footer', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 7 },
            { id: 8, type: 'CHAR', name: 'Footer Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 8 },
            { id: 9, type: 'CHAR', name: 'Footnote Text Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 9 },
            { id: 10, type: 'PARA', name: 'Header', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 10 },
            { id: 11, type: 'CHAR', name: 'Header Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 11 },
            { id: 12, type: 'PARA', name: 'Heading 1', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 12 },
            { id: 13, type: 'CHAR', name: 'Heading 1 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 13 },
            { id: 14, type: 'PARA', name: 'Heading 2', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 14 },
            { id: 15, type: 'CHAR', name: 'Heading 2 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 15 },
            { id: 16, type: 'PARA', name: 'Heading 3', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 16 },
            { id: 17, type: 'CHAR', name: 'Heading 3 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 17 },
            { id: 18, type: 'PARA', name: 'Heading 4', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 18 },
            { id: 19, type: 'CHAR', name: 'Heading 4 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 19 },
            { id: 20, type: 'PARA', name: 'Heading 5', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 20 },
            { id: 21, type: 'CHAR', name: 'Heading 5 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 21 },
            { id: 22, type: 'PARA', name: 'Heading 6', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 22 },
            { id: 23, type: 'CHAR', name: 'Heading 6 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 23 },
            { id: 24, type: 'PARA', name: 'Heading 7', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 24 },
            { id: 25, type: 'CHAR', name: 'Heading 7 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 25 },
            { id: 26, type: 'PARA', name: 'Heading 8', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 26 },
            { id: 27, type: 'CHAR', name: 'Heading 8 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 27 },
            { id: 28, type: 'PARA', name: 'Heading 9', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 28 },
            { id: 29, type: 'CHAR', name: 'Heading 9 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 29 },
            { id: 30, type: 'CHAR', name: 'Hyperlink', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 30 },
            { id: 31, type: 'CHAR', name: 'Intense Emphasis', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 31 },
            { id: 32, type: 'PARA', name: 'Intense Quote', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 32 },
            { id: 33, type: 'CHAR', name: 'Intense Quote Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 33 },
            { id: 34, type: 'CHAR', name: 'Intense Reference', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 34 },
            { id: 35, type: 'PARA', name: 'List Paragraph', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 35 },
            { id: 36, type: 'PARA', name: 'No Spacing', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 36 },
            { id: 37, type: 'CHAR', name: 'Placeholder Text', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 37 },
            { id: 38, type: 'PARA', name: 'Quote', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 38 },
            { id: 39, type: 'CHAR', name: 'Quote Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 39 },
            { id: 40, type: 'CHAR', name: 'Strong', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 40 },
            { id: 41, type: 'PARA', name: 'Subtitle', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 41 },
            { id: 42, type: 'CHAR', name: 'Subtitle Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 42 },
            { id: 43, type: 'CHAR', name: 'Subtle Emphasis', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 43 },
            { id: 44, type: 'CHAR', name: 'Subtle Reference', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 44 },
            { id: 45, type: 'PARA', name: 'TOC Heading', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 45 },
            { id: 46, type: 'PARA', name: 'Title', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 46 },
            { id: 47, type: 'CHAR', name: 'Title Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 47 },
            { id: 48, type: 'CHAR', name: 'endnote reference', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 48 },
            { id: 49, type: 'PARA', name: 'endnote text', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 49 },
            { id: 50, type: 'CHAR', name: 'footnote reference', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 50 },
            { id: 51, type: 'PARA', name: 'footnote text', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 51 },
            { id: 52, type: 'PARA', name: 'table of figures', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 52 },
            { id: 53, type: 'PARA', name: 'toc 1', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 53 },
            { id: 54, type: 'PARA', name: 'toc 2', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 54 },
            { id: 55, type: 'PARA', name: 'toc 3', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 55 },
            { id: 56, type: 'PARA', name: 'toc 4', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 56 },
            { id: 57, type: 'PARA', name: 'toc 5', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 57 },
            { id: 58, type: 'PARA', name: 'toc 6', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 58 },
            { id: 59, type: 'PARA', name: 'toc 7', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 59 },
            { id: 60, type: 'PARA', name: 'toc 8', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 60 },
            { id: 61, type: 'PARA', name: 'toc 9', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 61 },
        ];

        // DOCX 스타일 ID → HWPX 스타일 ID 매핑 테이블 저장 (다른 메서드에서 참조용)
        this.state.docxStyleToHwpxId = {
            'Normal': 0,
            'Heading1': 12, 'Heading2': 14, 'Heading3': 16,
            'Heading4': 18, 'Heading5': 20, 'Heading6': 22,
            'Heading7': 24, 'Heading8': 26, 'Heading9': 28,
            'Title': 46, 'Subtitle': 41,
            'ListParagraph': 35, 'NoSpacing': 36,
            'Quote': 38, 'IntenseQuote': 32,
            'Header': 10, 'Footer': 7,
            'FootnoteText': 51, 'EndnoteText': 49,
            'TOCHeading': 45,
            'TOC1': 53, 'TOC2': 54, 'TOC3': 55,
            'TOC4': 56, 'TOC5': 57, 'TOC6': 58,
            'TOC7': 59, 'TOC8': 60, 'TOC9': 61,
            'Caption': 2,
            'TableofFigures': 52,
        };

        let xml = `<hh:styles itemCnt="${styleDefinitions.length}">`;
        for (const s of styleDefinitions) {
            xml += `<hh:style id="${s.id}" type="${s.type}" name="${escapeXml(s.name)}" engName="${escapeXml(s.engName)}" paraPrIDRef="${s.paraPrIDRef}" charPrIDRef="${s.charPrIDRef}" nextStyleIDRef="${s.nextStyleIDRef}" langID="1042" lockForm="0"/>`;
        }
        xml += `</hh:styles>`;
        return xml;
    }


}
