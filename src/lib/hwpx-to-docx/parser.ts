import { ConversionContext } from './context';
import { BorderFill } from './types';

export class HwpxParser {
    private ctx: ConversionContext;

    constructor(ctx: ConversionContext) {
        this.ctx = ctx;
    }

    public async parseContentHpf(_xml: string) {
        // console.log('content.hpf 파싱 완료');
    }

    public parseHeader(xml: string) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'text/xml');

        this.parseFontFaces(doc);
        this.parseBorderFills(doc);
        this.parseCharProperties(doc);
        this.parseParaProperties(doc);
        this.parseStyles(doc);
        this.parseNumberings(doc);

        console.log(`header.xml 파싱 완료: charPr=${this.ctx.charProperties.size}, paraPr=${this.ctx.paraProperties.size}, borderFill=${this.ctx.borderFills.size}`);
    }

    private parseFontFaces(doc: Document) {
        const fontfaceElems = doc.getElementsByTagNameNS(this.ctx.ns.hh, 'fontface');
        for (const ff of Array.from(fontfaceElems)) {
            const lang = ff.getAttribute('lang') || 'UNKNOWN';
            const fonts: { id: string, face: string }[] = [];
            const fontElems = ff.getElementsByTagNameNS(this.ctx.ns.hh, 'font');
            for (const f of Array.from(fontElems)) {
                fonts.push({
                    id: f.getAttribute('id') || '',
                    face: f.getAttribute('face') || ''
                });
            }
            this.ctx.fontFaces[lang] = fonts;
        }
    }

    private parseBorderFills(doc: Document) {
        let borderFills: Element[] = [];

        // getElementsByTagNameNS가 잘 안될 경우를 대비해 태그명으로 직접 검색
        const allElements = doc.getElementsByTagName('*');
        for (let i = 0; i < allElements.length; i++) {
            if (allElements[i].localName === 'borderFill') {
                borderFills.push(allElements[i]);
            }
        }

        for (const bf of borderFills) {
            const id = bf.getAttribute('id') || '';
            const parseBorder = (name: string) => {
                let borderElem = bf.getElementsByTagNameNS(this.ctx.ns.hh, name)[0];
                if (!borderElem) {
                    const all = bf.getElementsByTagName(`hh:${name}`);
                    if (all.length > 0) borderElem = all[0];
                }
                if (borderElem) {
                    return {
                        type: borderElem.getAttribute('type') || 'none',
                        width: borderElem.getAttribute('width') || '0',
                        color: borderElem.getAttribute('color') || '#000000'
                    };
                }
                return { type: 'none', width: '0', color: '#000000' };
            };

            const borderFill: BorderFill = {
                id,
                leftBorder: parseBorder('leftBorder'),
                rightBorder: parseBorder('rightBorder'),
                topBorder: parseBorder('topBorder'),
                bottomBorder: parseBorder('bottomBorder')
            };

            // 채우기 파싱 (localName 기반 - namespace 불일치 방지)
            const fillBrushElem = Array.from(bf.childNodes)
                .find(n => n.nodeType === 1 && (n as Element).localName === 'fillBrush') as Element | undefined;
            if (fillBrushElem) {
                const winBrushElem = Array.from(fillBrushElem.childNodes)
                    .find(n => n.nodeType === 1 && (n as Element).localName === 'winBrush') as Element | undefined;
                if (winBrushElem) {
                    const fc = winBrushElem.getAttribute('faceColor');
                    if (fc && fc !== 'none') borderFill.faceColor = fc;
                }
            }

            this.ctx.borderFills.set(id, borderFill);
        }
    }

    private parseCharProperties(doc: Document) {
        const charPrs = doc.getElementsByTagName('hh:charPr');
        for (let i = 0; i < charPrs.length; i++) {
            const cp = charPrs[i];
            const id = cp.getAttribute('id') || '';
            const height = parseInt(cp.getAttribute('height') || '1000');
            const textColor = cp.getAttribute('textColor') || '#000000';
            const shadeColor = cp.getAttribute('shadeColor') || 'none';
            const borderFillIDRef = cp.getAttribute('borderFillIDRef') || '';

            // 자식 요소 헬퍼 (localName 기반으로 namespace 불일치 방지)
            const getChild = (localName: string) =>
                Array.from(cp.childNodes).find(
                    n => n.nodeType === 1 && (n as Element).localName === localName
                ) as Element | undefined;

            // bold, italic: <hh:bold value="true"/> 형태
            const bold = getChild('bold')?.getAttribute('value') === 'true';
            const italic = getChild('italic')?.getAttribute('value') === 'true';

            // underline: type이 NONE이 아니면 적용
            const underlineElem = getChild('underline');
            const underline = !!underlineElem && (underlineElem.getAttribute('type') || 'NONE') !== 'NONE';

            // strikeout: type이 NONE이 아니면 적용
            const strikeoutElem = getChild('strikeout');
            const strikeout = !!strikeoutElem && (strikeoutElem.getAttribute('type') || 'NONE') !== 'NONE';

            // 위/아래 첨자
            const supscript = getChild('superScript')?.getAttribute('value') === 'true';
            const subscript = getChild('subScript')?.getAttribute('value') === 'true';

            // 폰트: <hh:fontRef hangul="0" latin="0" .../>에서 ID 읽고 fontFaces 맵에서 이름 조회
            let hangulFont = '함초롬바탕', latinFont = '함초롬바탕';
            const fontRefElem = getChild('fontRef');
            if (fontRefElem) {
                const hangulId = fontRefElem.getAttribute('hangul') || '0';
                const latinId  = fontRefElem.getAttribute('latin')  || '0';

                const hangulFaces = this.ctx.fontFaces['HANGUL'] || this.ctx.fontFaces['Hangul'] || [];
                const latinFaces  = this.ctx.fontFaces['LATIN']  || this.ctx.fontFaces['Latin']  || [];

                const hf = hangulFaces.find(f => f.id === hangulId);
                const lf = latinFaces.find(f => f.id === latinId);
                if (hf?.face) hangulFont = hf.face;
                if (lf?.face) latinFont  = lf.face;
            }

            this.ctx.charProperties.set(id, {
                id, height, textColor, shadeColor, borderFillIDRef,
                bold, italic, underline, strikeout, supscript, subscript,
                hangulFont, latinFont
            });
        }
    }

    private parseParaProperties(doc: Document) {
        const paraPrs = doc.getElementsByTagName('hh:paraPr');
        for (let i = 0; i < paraPrs.length; i++) {
            const pp = paraPrs[i];
            const id = pp.getAttribute('id') || '';

            // 자식 요소 헬퍼
            const getChild = (localName: string) =>
                Array.from(pp.childNodes).find(
                    n => n.nodeType === 1 && (n as Element).localName === localName
                ) as Element | undefined;

            // <hh:margin left="..." right="..." prev="..." next="...">
            //   <hh:indent value="..."/>
            // </hh:margin>
            let leftMargin = 0, rightMargin = 0, indent = 0, prevSpacing = 0, nextSpacing = 0;
            const marginElem = getChild('margin');
            if (marginElem) {
                leftMargin  = parseInt(marginElem.getAttribute('left')  || '0', 10);
                rightMargin = parseInt(marginElem.getAttribute('right') || '0', 10);
                prevSpacing = parseInt(marginElem.getAttribute('prev')  || '0', 10);
                nextSpacing = parseInt(marginElem.getAttribute('next')  || '0', 10);

                const indentElem = Array.from(marginElem.childNodes).find(
                    n => n.nodeType === 1 && (n as Element).localName === 'indent'
                ) as Element | undefined;
                if (indentElem) {
                    indent = parseInt(indentElem.getAttribute('value') || '0', 10);
                }
            }

            // <hh:lineSpacing type="PERCENT" value="160"/>
            let lineSpacingType = 'PERCENT', lineSpacingVal = 160;
            const lsElem = getChild('lineSpacing');
            if (lsElem) {
                lineSpacingType = lsElem.getAttribute('type') || 'PERCENT';
                lineSpacingVal  = parseInt(lsElem.getAttribute('value') || '160', 10);
            }

            // <hh:align value="JUSTIFY"/>
            const alignElem = getChild('align');
            const align = alignElem?.getAttribute('value') || 'JUSTIFY';

            // keepWithNext, keepLines, pageBreakBefore (속성 또는 자식요소 양쪽 모두 처리)
            const keepWithNext = pp.getAttribute('keepWithNext') === 'true' ? '1' : '0';
            const keepLines    = pp.getAttribute('keepLines')    === 'true' ? '1' : '0';
            const pageBreakBefore = pp.getAttribute('pageBreakBefore') === 'true' ? '1' : '0';

            // <hh:heading type="NONE" idRef="0" level="0"/>
            const headingElem = getChild('heading');
            const heading     = headingElem?.getAttribute('type')  || 'NONE';
            const level       = headingElem?.getAttribute('level') || '0';
            const headingIdRef = headingElem?.getAttribute('idRef') || '0';

            // <hh:borderFill borderFillIDRef="0"/>
            const bfElem = getChild('borderFill');
            const borderFillIDRef = bfElem?.getAttribute('borderFillIDRef') || '';

            this.ctx.paraProperties.set(id, {
                id, align, heading, level, headingIdRef,
                leftMargin, rightMargin, indent, prevSpacing, nextSpacing,
                lineSpacingType, lineSpacingVal,
                keepWithNext, keepLines, pageBreakBefore, borderFillIDRef
            });
        }
    }

    private parseStyles(_doc: Document) {
        // ...
    }

    private parseNumberings(_doc: Document) {
        // ...
    }
}
