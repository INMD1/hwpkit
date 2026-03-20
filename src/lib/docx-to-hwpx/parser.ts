import JSZip from 'jszip';
import { ConversionState } from './state';
import { StyleData } from './types';
import { getImageDimensions } from './utils';

export class DocxParser {
    private state: ConversionState;

    constructor(state: ConversionState) {
        this.state = state;
    }

    public async parseMetadata(zip: JSZip) {
        try {
            const coreXml = await zip.file('docProps/core.xml')?.async('text');
            if (coreXml) {
                const parser = new DOMParser();
                const coreDoc = parser.parseFromString(coreXml, 'text/xml');
                const getTag = (tag: string) => coreDoc.getElementsByTagName(tag)[0]?.textContent || '';

                this.state.metadata.title = getTag('dc:title') || getTag('title');
                this.state.metadata.creator = getTag('dc:creator') || getTag('creator');
                this.state.metadata.subject = getTag('dc:subject') || getTag('subject');
                this.state.metadata.description = getTag('dc:description') || getTag('description');
                this.state.metadata.keywords = getTag('cp:keywords') || getTag('keywords');
                this.state.metadata.lastModifiedBy = getTag('cp:lastModifiedBy') || getTag('lastModifiedBy');
                this.state.metadata.revision = getTag('cp:revision') || getTag('revision');
                this.state.metadata.created = getTag('dcterms:created') || getTag('created');
                this.state.metadata.modified = getTag('dcterms:modified') || getTag('modified');
            }
        } catch (e) {
            console.warn('메타데이터 파싱 실패:', e);
        }
    }

    public async parseRelationships(zip: JSZip) {
        try {
            const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('text');
            if (relsXml) {
                const parser = new DOMParser();
                const relsDoc = parser.parseFromString(relsXml, 'text/xml');
                const rels = relsDoc.getElementsByTagName('Relationship');
                for (let i = 0; i < rels.length; i++) {
                    const rel = rels[i];
                    const id = rel.getAttribute('Id');
                    const target = rel.getAttribute('Target');
                    const type = rel.getAttribute('Type');
                    if (id && target && type) {
                        this.state.relationships.set(id, { target, type });
                    }
                }
            }
        } catch (e) {
            console.warn('관계 파싱 실패:', e);
        }
    }

    public async parseImages(zip: JSZip) {
        for (const [rId, rel] of this.state.relationships) {
            if (rel.type?.includes('image')) {
                try {
                    let targetPath = rel.target;
                    if (targetPath.startsWith('/')) targetPath = targetPath.slice(1);
                    const imagePath = targetPath.startsWith('word/') ? targetPath : `word/${targetPath}`;
                    const imageFile = zip.file(imagePath);
                    if (imageFile) {
                        const imageData = await imageFile.async('uint8array');
                        const ext = (imagePath.split('.').pop() || '').toLowerCase();
                        const manifestId = `BIN${String(this.state.imageCounter).padStart(4, '0')}`;

                        // 이미지 크기 추출
                        const dimensions = getImageDimensions(imageData, ext);

                        this.state.images.set(rId, {
                            data: imageData,
                            ext: ext === 'jpeg' ? 'jpg' : ext, // Standardize extension
                            manifestId: manifestId,
                            path: `BinData/${manifestId}.${ext === 'jpeg' ? 'jpg' : ext}`,
                            width: dimensions.width,
                            height: dimensions.height,
                            binDataId: this.state.imageCounter
                        });
                        this.state.imageCounter++;
                    }
                } catch (e) {
                    console.warn(`이미지 ${rId} 파싱 실패:`, e);
                }
            }
        }
        console.log(`✅ ${this.state.images.size}개 이미지 파싱 완료`);
    }

    public async parseStyles(zip: JSZip) {
        try {
            const stylesXml = await zip.file('word/styles.xml')?.async('text');
            if (stylesXml) {
                const parser = new DOMParser();
                const stylesDoc = parser.parseFromString(stylesXml, 'text/xml');
                const styles = stylesDoc.getElementsByTagNameNS(this.state.ns.w, 'style');

                for (let i = 0; i < styles.length; i++) {
                    const style = styles[i];
                    const styleId = style.getAttribute('w:styleId');
                    // const type = style.getAttribute('w:type');

                    if (styleId) {
                        let styleData: StyleData = this.state.docStyles.get(styleId) || {};

                        // Paragraph Properties within Style
                        const pPr = style.getElementsByTagNameNS(this.state.ns.w, 'pPr')[0];
                        if (pPr) {
                            // Alignment
                            const jc = pPr.getElementsByTagNameNS(this.state.ns.w, 'jc')[0];
                            if (jc) styleData.align = jc.getAttribute('w:val') || undefined;

                            // Indentation
                            const ind = pPr.getElementsByTagNameNS(this.state.ns.w, 'ind')[0];
                            if (ind) {
                                styleData.ind = {};
                                if (ind.getAttribute('w:left')) styleData.ind.left = ind.getAttribute('w:left') || undefined;
                                if (ind.getAttribute('w:right')) styleData.ind.right = ind.getAttribute('w:right') || undefined;
                                if (ind.getAttribute('w:firstLine')) styleData.ind.firstLine = ind.getAttribute('w:firstLine') || undefined;
                                if (ind.getAttribute('w:hanging')) styleData.ind.hanging = ind.getAttribute('w:hanging') || undefined;
                            }

                            // Spacing
                            const spacing = pPr.getElementsByTagNameNS(this.state.ns.w, 'spacing')[0];
                            if (spacing) {
                                styleData.spacing = {};
                                if (spacing.getAttribute('w:before')) styleData.spacing.before = spacing.getAttribute('w:before') || undefined;
                                if (spacing.getAttribute('w:after')) styleData.spacing.after = spacing.getAttribute('w:after') || undefined;
                                if (spacing.getAttribute('w:line')) styleData.spacing.line = spacing.getAttribute('w:line') || undefined;
                                if (spacing.getAttribute('w:lineRule')) styleData.spacing.lineRule = spacing.getAttribute('w:lineRule') || undefined;
                            }
                        }

                        // Run Properties within Style
                        const rPr = style.getElementsByTagNameNS(this.state.ns.w, 'rPr')[0];
                        if (rPr) {
                            if (!styleData.rPr) styleData.rPr = {};

                            const sz = rPr.getElementsByTagNameNS(this.state.ns.w, 'sz')[0];
                            if (sz) styleData.rPr.sz = sz.getAttribute('w:val') || undefined;

                            const color = rPr.getElementsByTagNameNS(this.state.ns.w, 'color')[0];
                            if (color) styleData.rPr.color = color.getAttribute('w:val') || undefined;
                        }

                        // Table Properties
                        const tblPr = style.getElementsByTagNameNS(this.state.ns.w, 'tblPr')[0];
                        if (tblPr) {
                            const tblBorders = tblPr.getElementsByTagNameNS(this.state.ns.w, 'tblBorders')[0];
                            if (tblBorders) styleData.tblBorders = tblBorders;
                        }

                        // Table Cell Properties
                        const tcPr = style.getElementsByTagNameNS(this.state.ns.w, 'tcPr')[0];
                        if (tcPr) {
                            const tcBorders = tcPr.getElementsByTagNameNS(this.state.ns.w, 'tcBorders')[0];
                            if (tcBorders) styleData.tcBorders = tcBorders;
                            const shd = tcPr.getElementsByTagNameNS(this.state.ns.w, 'shd')[0];
                            if (shd) styleData.shd = shd;
                        }

                        // Based On (Inheritance)
                        const basedOn = style.getElementsByTagNameNS(this.state.ns.w, 'basedOn')[0];
                        if (basedOn && !styleData.basedOn) styleData.basedOn = basedOn.getAttribute('w:val') || undefined;

                        this.state.docStyles.set(styleId, styleData);
                    }
                }
            }
        } catch (e) {
            console.warn('스타일 파싱 실패:', e);
        }
    }

    public async parseNumbering(zip: JSZip) {
        try {
            const numXml = await zip.file('word/numbering.xml')?.async('text');
            if (numXml) {
                const parser = new DOMParser();
                const numDoc = parser.parseFromString(numXml, 'text/xml');

                // Abstract Numbering
                const abstractNums = numDoc.getElementsByTagNameNS(this.state.ns.w, 'abstractNum');
                for (let i = 0; i < abstractNums.length; i++) {
                    const abs = abstractNums[i];
                    const absId = abs.getAttribute('w:abstractNumId');
                    if (absId) {
                        const levels = new Map();
                        const lvls = abs.getElementsByTagNameNS(this.state.ns.w, 'lvl');
                        for (let j = 0; j < lvls.length; j++) {
                            const lvl = lvls[j];
                            const ilvl = lvl.getAttribute('w:ilvl');
                            const start = parseInt(lvl.getElementsByTagNameNS(this.state.ns.w, 'start')[0]?.getAttribute('w:val') || '1');
                            const numFmt = lvl.getElementsByTagNameNS(this.state.ns.w, 'numFmt')[0]?.getAttribute('w:val') || 'decimal';
                            const lvlText = lvl.getElementsByTagNameNS(this.state.ns.w, 'lvlText')[0]?.getAttribute('w:val') || '';

                            // 레벨별 들여쓰기 정보 파싱
                            let lvlLeft = null;
                            let lvlHanging = null;
                            const lvlPPr = lvl.getElementsByTagNameNS(this.state.ns.w, 'pPr')[0];
                            if (lvlPPr) {
                                const lvlInd = lvlPPr.getElementsByTagNameNS(this.state.ns.w, 'ind')[0];
                                if (lvlInd) {
                                    lvlLeft = lvlInd.getAttribute('w:left') || lvlInd.getAttribute('w:start');
                                    lvlHanging = lvlInd.getAttribute('w:hanging');
                                }
                            }

                            if (ilvl) {
                                levels.set(ilvl, { start, numFmt, lvlText, lvlLeft, lvlHanging });
                            }
                        }
                        this.state.abstractNumMap.set(absId, levels);
                    }
                }

                // Numbering Instances
                const nums = numDoc.getElementsByTagNameNS(this.state.ns.w, 'num');
                for (let i = 0; i < nums.length; i++) {
                    const num = nums[i];
                    const numId = num.getAttribute('w:numId');
                    const absRef = num.getElementsByTagNameNS(this.state.ns.w, 'abstractNumId')[0]?.getAttribute('w:val');
                    if (numId && absRef) {
                        this.state.numberingMap.set(numId, absRef);
                    }
                }
            }
        } catch (e) {
            console.warn('번호 매기기 파싱 실패:', e);
        }
    }

    public async parseFootnotes(zip: JSZip) {
        try {
            const xml = await zip.file('word/footnotes.xml')?.async('text');
            if (xml) {
                const parser = new DOMParser();
                const doc = parser.parseFromString(xml, 'text/xml');
                const footnotes = doc.getElementsByTagNameNS(this.state.ns.w, 'footnote');
                for (let i = 0; i < footnotes.length; i++) {
                    const fn = footnotes[i];
                    const id = fn.getAttribute('w:id');
                    if (id) {
                        this.state.footnoteMap.set(id, fn);
                    }
                }
            }
        } catch (e) {
            console.warn('각주 파싱 실패:', e);
        }
    }

    public async parseEndnotes(zip: JSZip) {
        try {
            const xml = await zip.file('word/endnotes.xml')?.async('text');
            if (xml) {
                const parser = new DOMParser();
                const doc = parser.parseFromString(xml, 'text/xml');
                const endnotes = doc.getElementsByTagNameNS(this.state.ns.w, 'endnote');
                for (let i = 0; i < endnotes.length; i++) {
                    const en = endnotes[i];
                    const id = en.getAttribute('w:id');
                    if (id) {
                        this.state.endnoteMap.set(id, en);
                    }
                }
            }
        } catch (e) {
            console.warn('미주 파싱 실패:', e);
        }
    }

    public parsePageSetup(doc: Document) {
        const sectPr = doc.getElementsByTagNameNS(this.state.ns.w, 'sectPr')[0];
        if (!sectPr) return;

        // Page Size
        const pgSz = sectPr.getElementsByTagNameNS(this.state.ns.w, 'pgSz')[0];
        if (pgSz) {
            const w = parseInt(pgSz.getAttribute('w:w') || '11906'); // Twips (default A4)
            const h = parseInt(pgSz.getAttribute('w:h') || '16838');

            // Twips to HWPUNIT (1 Twip = 5 HWPUNIT)
            this.state.PAGE_WIDTH = w * 5;
            this.state.PAGE_HEIGHT = h * 5;
            this.state.TEXT_WIDTH = this.state.PAGE_WIDTH - this.state.MARGIN_LEFT - this.state.MARGIN_RIGHT; // Update text width
        }

        // Page Margins
        const pgMar = sectPr.getElementsByTagNameNS(this.state.ns.w, 'pgMar')[0];
        if (pgMar) {
            const top = parseInt(pgMar.getAttribute('w:top') || '1440');
            const bottom = parseInt(pgMar.getAttribute('w:bottom') || '1440');
            const left = parseInt(pgMar.getAttribute('w:left') || '1800');
            const right = parseInt(pgMar.getAttribute('w:right') || '1800');
            const header = parseInt(pgMar.getAttribute('w:header') || '720'); // Header margin
            const footer = parseInt(pgMar.getAttribute('w:footer') || '720'); // Footer margin
            const gutter = parseInt(pgMar.getAttribute('w:gutter') || '0');

            // Twips to HWPUNIT
            this.state.MARGIN_TOP = top * 5;
            this.state.MARGIN_BOTTOM = bottom * 5;
            this.state.MARGIN_LEFT = left * 5;
            this.state.MARGIN_RIGHT = right * 5;
            this.state.HEADER_MARGIN = header * 5;
            this.state.FOOTER_MARGIN = footer * 5;
            this.state.GUTTER_MARGIN = gutter * 5;

            // Update Text Width based on new margins
            this.state.TEXT_WIDTH = this.state.PAGE_WIDTH - this.state.MARGIN_LEFT - this.state.MARGIN_RIGHT - this.state.GUTTER_MARGIN;
        }

        // Columns
        const cols = sectPr.getElementsByTagNameNS(this.state.ns.w, 'cols')[0];
        if (cols) {
            this.state.COL_COUNT = parseInt(cols.getAttribute('w:num') || '1');
            const space = parseInt(cols.getAttribute('w:space') || '720');
            this.state.COL_GAP = space * 5;
            this.state.COL_TYPE = this.state.COL_COUNT > 1 ? 'NEWSPAPER' : 'ONE';
        }
    }
}
