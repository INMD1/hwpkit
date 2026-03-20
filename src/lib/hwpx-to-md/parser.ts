import { MdConversionContext } from './context';
import { BorderFill, BorderLine } from './types';

export class HwpxMdParser {
    private ctx: MdConversionContext;

    constructor(ctx: MdConversionContext) {
        this.ctx = ctx;
    }

    public parseHeader(xml: string) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, 'text/xml');

        this.parseCharProperties(doc);
        this.parseParaProperties(doc);
        this.parseStyles(doc);
        this.parseBorderFills(doc);
    }

    private parseCharProperties(doc: Document) {
        const charPrs = doc.getElementsByTagName('hh:charPr');
        for (let i = 0; i < charPrs.length; i++) {
            const cp = charPrs[i];
            const id = cp.getAttribute('id') || '';
            const height = parseInt(cp.getAttribute('height') || '1000');

            // <hh:bold hangul="1" latin="1" ...>
            const boldElem = cp.getElementsByTagNameNS(this.ctx.ns.hh, 'bold')[0];
            const bold = boldElem
                ? boldElem.getAttribute('hangul') === '1' || boldElem.getAttribute('latin') === '1'
                : false;

            // <hh:italic hangul="1" latin="1" ...>
            const italicElem = cp.getElementsByTagNameNS(this.ctx.ns.hh, 'italic')[0];
            const italic = italicElem
                ? italicElem.getAttribute('hangul') === '1' || italicElem.getAttribute('latin') === '1'
                : false;

            // <hh:underline type="SINGLE" ...>
            const underlineElem = cp.getElementsByTagNameNS(this.ctx.ns.hh, 'underline')[0];
            const underlineType = underlineElem?.getAttribute('type') ?? 'NONE';
            const underline = underlineType !== 'NONE';

            // <hh:strikeout type="SINGLE" ...>
            const strikeoutElem = cp.getElementsByTagNameNS(this.ctx.ns.hh, 'strikeout')[0];
            const strikeoutType = strikeoutElem?.getAttribute('type') ?? 'NONE';
            const strikeout = strikeoutType !== 'NONE';

            // <hh:superScript type="SUPERSCRIPT" ...>
            const superElem = cp.getElementsByTagNameNS(this.ctx.ns.hh, 'superScript')[0];
            const superType = superElem?.getAttribute('type') ?? 'NONE';
            const supscript = superType !== 'NONE';

            // <hh:subScript type="SUBSCRIPT" ...>
            const subElem = cp.getElementsByTagNameNS(this.ctx.ns.hh, 'subScript')[0];
            const subType = subElem?.getAttribute('type') ?? 'NONE';
            const subscript = subType !== 'NONE';

            this.ctx.charProperties.set(id, {
                id, height, bold, italic, underline, strikeout, supscript, subscript,
            });
        }
    }

    private parseParaProperties(doc: Document) {
        const paraPrs = doc.getElementsByTagName('hh:paraPr');
        for (let i = 0; i < paraPrs.length; i++) {
            const pp = paraPrs[i];
            const id = pp.getAttribute('id') || '';
            const align = pp.getAttribute('align') || 'LEFT';
            this.ctx.paraProperties.set(id, { id, align });
        }
    }

    private parseStyles(doc: Document) {
        const styles = doc.getElementsByTagName('hh:style');
        for (let i = 0; i < styles.length; i++) {
            const style = styles[i];
            const id = style.getAttribute('id') || '';
            const type = style.getAttribute('type') || 'PARA';
            const level = parseInt(style.getAttribute('level') || '0');
            const paraPrIDRef = style.getAttribute('paraPrIDRef') || '';
            const charPrIDRef = style.getAttribute('charPrIDRef') || '';
            this.ctx.styles.set(id, { id, type, level, paraPrIDRef, charPrIDRef });
        }
    }

    private parseBorderFills(doc: Document) {
        const list = Array.from(
            doc.getElementsByTagNameNS(this.ctx.ns.hh, 'borderFill')
        );

        for (const bf of list) {
            const id = bf.getAttribute('id');
            if (!id) continue;

            const parseLine = (tag: string): BorderLine | undefined => {
                const el = bf.getElementsByTagNameNS(this.ctx.ns.hh, tag)[0];
                if (!el) return undefined;
                const width = parseInt(el.getAttribute('width') || '0', 10);
                const colorVal = el.getAttribute('color') || '0';
                const color = '#' + Number(colorVal).toString(16).padStart(6, '0');
                return { width, color };
            };

            const fill: BorderFill = {
                left: parseLine('leftBorder'),
                right: parseLine('rightBorder'),
                top: parseLine('topBorder'),
                bottom: parseLine('bottomBorder'),
            };

            // 배경색
            const back = bf.getElementsByTagNameNS(this.ctx.ns.hh, 'fillBrush')[0];
            if (back) {
                const colorEl = back.getElementsByTagNameNS(this.ctx.ns.hh, 'winBrush')[0];
                if (colorEl) {
                    const c = colorEl.getAttribute('faceColor');
                    if (c) fill.backgroundColor =
                        '#' + Number(c).toString(16).padStart(6, '0');
                }
            }

            this.ctx.borderFills.set(id, fill);
        }
    }

}
