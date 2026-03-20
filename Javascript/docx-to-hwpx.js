// ```
//                 GNU LESSER GENERAL PUBLIC LICENSE
//                    Version 2.1, February 1999
// ```

// Copyright (C) 2026 INMD1

// Everyone is permitted to copy and distribute verbatim copies
// of this license document, but changing it is not allowed.

// [This LGPL license applies to the project **ANY2HWPX** developed by **INMD1**.]

// ---

// ### Preamble

// The licenses for most software are designed to take away your freedom to share and change it. By contrast, the GNU Lesser General Public Licenses are intended to guarantee your freedom to share and change free software—to make sure the software is free for all its users.

// This license, the Lesser General Public License, applies to some specially designated software packages—typically libraries—of the Free Software Foundation and other authors who decide to use it. You can use it too.

// When we speak of free software, we are referring to freedom of use, not price. Our General Public Licenses are designed to make sure that you have the freedom to distribute copies of free software (and charge for this service if you wish), that you receive source code or can get it if you want it, that you can change the software and use pieces of it in new free programs, and that you are informed that you can do these things.

// To protect your rights, we need to make restrictions that forbid distributors to deny you these rights or to ask you to surrender these rights. These restrictions translate to certain responsibilities for you if you distribute copies of the library or if you modify it.

// For example, if you distribute copies of the library, whether gratis or for a fee, you must give the recipients all the rights that we gave you. You must make sure that they, too, receive or can get the source code. If you link other code with the library, you must provide complete object files so that users can relink them with the library after making changes.

// We protect your rights with a two-step method:
// (1) we copyright the library, and
// (2) we offer you this license, which gives you legal permission to copy, distribute and/or modify the library.

// ---

// ## TERMS AND CONDITIONS FOR COPYING, DISTRIBUTION AND MODIFICATION

// **0. Definitions.**
// This License Agreement applies to any software library or other program which contains a notice placed by the copyright holder saying it may be distributed under the terms of this Lesser General Public License.

// "Library" refers to any such software library or work.
// "Work based on the Library" means either the Library or any derivative work under copyright law.

// ---

// **1. You may copy and distribute verbatim copies of the Library's complete source code** as you receive it, in any medium, provided that you conspicuously and appropriately publish on each copy an appropriate copyright notice and disclaimer of warranty.

// ---

// **2. You may modify your copy or copies of the Library** or any portion of it, thus forming a work based on the Library, and copy and distribute such modifications provided that you also meet all of these conditions:

// * a) The modified work must itself be a software library.
// * b) You must cause the files modified to carry prominent notices stating that you changed the files and the date of any change.
// * c) You must cause the whole of the work to be licensed at no charge to all third parties under the terms of this License.
// * d) If a facility refers to a function or table supplied by an application program, the facility must still operate if the application does not supply it.

// ---

// **3. You may opt to apply the terms of the ordinary GNU General Public License** instead of this License to a given copy of the Library.

// ---

// **4. You may copy and distribute the Library (or a portion or derivative of it) in object code or executable form** under the terms of Sections 1 and 2 provided that you accompany it with the complete corresponding machine-readable source code.

// ---

// **5. A program that contains no derivative of any portion of the Library** but is designed to work with the Library by being compiled or linked with it is called a "work that uses the Library". Such a work is not a derivative work of the Library, and therefore falls outside the scope of this License.

// ---

// **6. As an exception**, you may combine or link a "work that uses the Library" with the Library to produce a work containing portions of the Library, and distribute that work under terms of your choice, provided that the terms permit modification of the work for the customer’s own use and reverse engineering for debugging such modifications.

// ---

// **7. You may place library facilities that are a work based on the Library side-by-side in a single library together with other library facilities not covered by this License**, and distribute such a combined library.

// ---

// **8. You may not copy, modify, sublicense, link with, or distribute the Library except as expressly provided under this License.** Any attempt otherwise is void and will automatically terminate your rights under this License.

// ---

// **9. You are not required to accept this License**, since you have not signed it. However, nothing else grants you permission to modify or distribute the Library or its derivative works.

// ---

// **10. Each time you redistribute the Library**, the recipient automatically receives a license from the original licensor to copy, distribute, link with or modify the Library subject to these terms.

// ---

// **11. If, as a consequence of a court judgment or allegation of patent infringement**, conditions are imposed on you that contradict the conditions of this License, they do not excuse you from the conditions of this License.

// ---

// **12. If the distribution and/or use of the Library is restricted in certain countries** either by patents or by copyrighted interfaces, the original copyright holder may add an explicit geographical distribution limitation.

// ---

// **13. The Free Software Foundation may publish revised and/or new versions** of the Lesser General Public License from time to time.

// ---

// **14. If you wish to incorporate parts of the Library into other free programs whose distribution conditions are incompatible**, write to the author to ask for permission.

// ---

// ## NO WARRANTY

// **15. BECAUSE THE LIBRARY IS LICENSED FREE OF CHARGE, THERE IS NO WARRANTY** for the Library, to the extent permitted by applicable law.

// ---

// **16. IN NO EVENT UNLESS REQUIRED BY APPLICABLE LAW OR AGREED TO IN WRITING** will any copyright holder be liable for damages arising out of the use or inability to use the Library.

// ---

// ## END OF TERMS AND CONDITIONS

// ---

// ### How to Apply These Terms to Your New Libraries

// To apply these terms to your own library, attach the following notices:

// ```
// ANY2HWPX — Copyright (C) 2026 INMD1

// This library is free software; you can redistribute it and/or
// modify it under the terms of the GNU Lesser General Public
// License as published by the Free Software Foundation; either
// version 2.1 of the License, or (at your option) any later version.

// This library is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
// See the GNU Lesser General Public License for more details.
// ```

// ---

class DocxToHwpxConverter {
    constructor() {
        this.ns = {
            w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
            pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
        };

        this.hwpxNs = {
            ha: 'http://www.hancom.co.kr/hwpml/2011/app',
            hp: 'http://www.hancom.co.kr/hwpml/2011/paragraph',
            hp10: 'http://www.hancom.co.kr/hwpml/2016/paragraph',
            hs: 'http://www.hancom.co.kr/hwpml/2011/section',
            hc: 'http://www.hancom.co.kr/hwpml/2011/core',
            hh: 'http://www.hancom.co.kr/hwpml/2011/head',
            hhs: 'http://www.hancom.co.kr/hwpml/2011/history',
            hm: 'http://www.hancom.co.kr/hwpml/2011/master-page',
            hpf: 'http://www.hancom.co.kr/schema/2011/hpf',
            dc: 'http://purl.org/dc/elements/1.1/',
            opf: 'http://www.idpf.org/2007/opf/',
            ooxmlchart: 'http://www.hancom.co.kr/hwpml/2016/ooxmlchart',
            epub: 'http://www.idpf.org/2007/ops',
            config: 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
            hv: 'http://www.hancom.co.kr/hwpml/2011/version',
            ocf: 'urn:oasis:names:tc:opendocument:xmlns:container',
            odf: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
            rdf: 'http://www.w3.org/1999/02/22-rdf-syntax-ns#',
            odfPkg: 'http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document'
        };

        // HWPUNIT = 1/7200 inch (PDF 표 1 참고)
        this.HWPUNIT_PER_INCH = 7200;
        this.EMU_PER_INCH = 914400; // EMU (English Metric Unit) for DOCX

        // 페이지 크기 (A4, PDF 표 131 참고)
        this.PAGE_WIDTH = 59530;  // A4 가로 (HWPUNIT)
        this.PAGE_HEIGHT = 84190; // A4 세로 (HWPUNIT)
        this.MARGIN_LEFT = 8505;  // 왼쪽 여백
        this.MARGIN_RIGHT = 8505; // 오른쪽 여백 (Default updated)
        this.MARGIN_TOP = 5668;   // 위 여백 (Default updated)
        this.MARGIN_BOTTOM = 4252; // 아래 여백 (Default updated)
        this.HEADER_MARGIN = 4252; // 머리말 여백
        this.FOOTER_MARGIN = 4252; // 꼬리말 여백
        this.GUTTER_MARGIN = 0;   // 제본 여백

        // 다단 설정
        this.COL_COUNT = 1;
        this.COL_GAP = 0;
        this.COL_TYPE = 'ONE';

        // 텍스트 영역 크기 계산
        this.TEXT_WIDTH = this.PAGE_WIDTH - this.MARGIN_LEFT - this.MARGIN_RIGHT - this.GUTTER_MARGIN;

        // Shape Table
        this.charShapes = new Map();
        this.paraShapes = new Map();
        this.borderFills = new Map();
        this.styles = new Map();

        // 카운터
        this.idCounter = 3121190098;
        this.tableIdCounter = 2085132891;
        this.imageCounter = 1;

        // 문단 위치 추적 (PDF 표 62 참고)
        this.currentVertPos = 0;

        // 각주/미주 순차 번호
        this.footnoteNumber = 0;
        this.endnoteNumber = 0;

        // 메타데이터
        this.metadata = {
            title: '', creator: '', subject: '', description: ''
        };

        // 자원
        this.relationships = new Map();
        this.images = new Map();

        // 언어별 폰트 추적 (HWPX는 각 언어마다 별도 폰트 리스트를 가짐)
        this.langFontFaces = {
            HANGUL: new Map(),
            LATIN: new Map(),
            HANJA: new Map(),
            JAPANESE: new Map(),
            OTHER: new Map(),
            SYMBOL: new Map(),
            USER: new Map()
        };
        // 기본 폰트 등록 (참조 형식)
        this.registerFontForLang('HANGUL', '함초롬바탕');
        this.registerFontForLang('LATIN', '함초롬바탕');
        this.registerFontForLang('HANJA', '함초롬바탕');
        this.registerFontForLang('JAPANESE', '함초롬바탕');
        this.registerFontForLang('OTHER', '함초롬바탕');
        this.registerFontForLang('SYMBOL', '함초롬바탕');
        this.registerFontForLang('USER', '함초롬바탕');
        // 호환용 (기존 fontFaces 참조가 있는 코드와의 호환)
        this.fontFaces = this.langFontFaces.HANGUL;

        // 번호 매기기 (Numbering)
        this.numberingMap = new Map(); // numId -> abstractNumId
        this.abstractNumMap = new Map(); // abstractNumId -> levels
        this.listCounters = new Map(); // "numId:ilvl" -> 현재 카운터 값
        this.listPrevNumId = null; // 이전 목록 numId 추적

        // 문서 스타일 (Styles)
        this.docStyles = new Map(); // styleId -> { align, ind, spacing, ... }

        // 주석 (Footnotes/Endnotes)
        this.footnoteMap = new Map();
        this.endnoteMap = new Map();

        // 첫 문단 플래그
        this.isFirstParagraph = true;
    }

    async convert(docxFile) {
        console.log('[1/8] DOCX 로드 중...');
        const zip = await JSZip.loadAsync(docxFile);

        console.log('[2/8] 메타데이터 파싱...');
        await this.parseMetadata(zip);

        console.log('[3/8] 이미지 및 관계 파싱...');
        await this.parseRelationships(zip);
        await this.parseImages(zip);
        await this.parseStyles(zip);
        await this.parseNumbering(zip);
        await this.parseFootnotes(zip);
        await this.parseEndnotes(zip);

        console.log('[4/8] DOCX 파싱 중...');
        const docXml = await zip.file('word/document.xml').async('text');
        const doc = new DOMParser().parseFromString(docXml, 'text/xml');

        // Page Setup Parsing
        this.parsePageSetup(doc);

        console.log('[5/8] 스타일 초기화...');
        this.initDefaultStyles();

        console.log('[6/8] HWPX 생성 중...');
        const hwpxData = this.createHwpx(doc);

        console.log('[7/8] 패키징 중...');
        const blob = await this.packageHwpx(hwpxData);

        console.log('[8/8] 완료!');
        return blob;
    }

    async parseMetadata(zip) {
        try {
            const coreXml = await zip.file('docProps/core.xml')?.async('text');
            if (coreXml) {
                const coreDoc = new DOMParser().parseFromString(coreXml, 'text/xml');
                const getTag = (tag) => coreDoc.getElementsByTagName(tag)[0]?.textContent || '';

                this.metadata.title = getTag('dc:title') || getTag('title');
                this.metadata.creator = getTag('dc:creator') || getTag('creator');
                this.metadata.subject = getTag('dc:subject') || getTag('subject');
                this.metadata.description = getTag('dc:description') || getTag('description');
                this.metadata.keywords = getTag('cp:keywords') || getTag('keywords');
                this.metadata.lastModifiedBy = getTag('cp:lastModifiedBy') || getTag('lastModifiedBy');
                this.metadata.revision = getTag('cp:revision') || getTag('revision');
                this.metadata.created = getTag('dcterms:created') || getTag('created');
                this.metadata.modified = getTag('dcterms:modified') || getTag('modified');
            }
        } catch (e) {
            console.warn('메타데이터 파싱 실패:', e);
        }
    }

    async parseRelationships(zip) {
        try {
            const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('text');
            if (relsXml) {
                const relsDoc = new DOMParser().parseFromString(relsXml, 'text/xml');
                const rels = relsDoc.getElementsByTagName('Relationship');
                for (const rel of rels) {
                    const id = rel.getAttribute('Id');
                    const target = rel.getAttribute('Target');
                    const type = rel.getAttribute('Type');
                    this.relationships.set(id, { target, type });
                }
            }
        } catch (e) {
            console.warn('관계 파싱 실패:', e);
        }
    }

    parsePageSetup(doc) {
        try {
            const body = doc.getElementsByTagNameNS(this.ns.w, 'body')[0];
            if (!body) return;

            // 마지막 섹션 속성 (전체 문서 또는 마지막 섹션의 설정)
            const sectPr = body.getElementsByTagNameNS(this.ns.w, 'sectPr')[0];
            if (sectPr) {
                this.extractSectionProperties(sectPr);
            }
        } catch (e) {
            console.warn('페이지 설정 파싱 실패:', e);
        }
    }

    extractSectionProperties(sectPr) {
        // Page Size
        const pgSz = sectPr.getElementsByTagNameNS(this.ns.w, 'pgSz')[0];
        if (pgSz) {
            const w = parseInt(pgSz.getAttribute('w:w') || '11906'); // Twips (default A4)
            const h = parseInt(pgSz.getAttribute('w:h') || '16838');
            const orient = pgSz.getAttribute('w:orient') || 'portrait';

            // Twips to HWPUNIT (1 Twip = 5 HWPUNIT)
            this.PAGE_WIDTH = w * 5;
            this.PAGE_HEIGHT = h * 5;
            this.TEXT_WIDTH = this.PAGE_WIDTH - this.MARGIN_LEFT - this.MARGIN_RIGHT; // Update text width
        }

        // Page Margins
        const pgMar = sectPr.getElementsByTagNameNS(this.ns.w, 'pgMar')[0];
        if (pgMar) {
            const top = parseInt(pgMar.getAttribute('w:top') || '1440');
            const bottom = parseInt(pgMar.getAttribute('w:bottom') || '1440');
            const left = parseInt(pgMar.getAttribute('w:left') || '1800');
            const right = parseInt(pgMar.getAttribute('w:right') || '1800');
            const header = parseInt(pgMar.getAttribute('w:header') || '720'); // Header margin
            const footer = parseInt(pgMar.getAttribute('w:footer') || '720'); // Footer margin
            const gutter = parseInt(pgMar.getAttribute('w:gutter') || '0');

            // Twips to HWPUNIT
            this.MARGIN_TOP = top * 5;
            this.MARGIN_BOTTOM = bottom * 5;
            this.MARGIN_LEFT = left * 5;
            this.MARGIN_RIGHT = right * 5;
            this.HEADER_MARGIN = header * 5;
            this.FOOTER_MARGIN = footer * 5;
            this.GUTTER_MARGIN = gutter * 5;

            // Update Text Width based on new margins
            this.TEXT_WIDTH = this.PAGE_WIDTH - this.MARGIN_LEFT - this.MARGIN_RIGHT - this.GUTTER_MARGIN;
        }

        // Columns
        const cols = sectPr.getElementsByTagNameNS(this.ns.w, 'cols')[0];
        if (cols) {
            this.COL_COUNT = parseInt(cols.getAttribute('w:num') || '1');
            const space = parseInt(cols.getAttribute('w:space') || '720');
            this.COL_GAP = space * 5;
            this.COL_TYPE = this.COL_COUNT > 1 ? 'NEWSPAPER' : 'ONE';
        }
    }

    async parseStyles(zip) {
        try {
            const stylesXml = await zip.file('word/styles.xml')?.async('text');
            if (stylesXml) {
                const stylesDoc = new DOMParser().parseFromString(stylesXml, 'text/xml');
                const styles = stylesDoc.getElementsByTagNameNS(this.ns.w, 'style');

                for (const style of styles) {
                    const styleId = style.getAttribute('w:styleId');
                    const type = style.getAttribute('w:type');

                    if (type === 'paragraph') {
                        const styleData = {};

                        // Paragraph Properties within Style
                        const pPr = style.getElementsByTagNameNS(this.ns.w, 'pPr')[0];
                        if (pPr) {
                            // Alignment
                            const jc = pPr.getElementsByTagNameNS(this.ns.w, 'jc')[0];
                            if (jc) styleData.align = jc.getAttribute('w:val');

                            // Indentation
                            const ind = pPr.getElementsByTagNameNS(this.ns.w, 'ind')[0];
                            if (ind) {
                                styleData.ind = {};
                                if (ind.getAttribute('w:left')) styleData.ind.left = ind.getAttribute('w:left');
                                if (ind.getAttribute('w:right')) styleData.ind.right = ind.getAttribute('w:right');
                                if (ind.getAttribute('w:firstLine')) styleData.ind.firstLine = ind.getAttribute('w:firstLine');
                                if (ind.getAttribute('w:hanging')) styleData.ind.hanging = ind.getAttribute('w:hanging');
                            }

                            // Spacing
                            const spacing = pPr.getElementsByTagNameNS(this.ns.w, 'spacing')[0];
                            if (spacing) {
                                styleData.spacing = {};
                                if (spacing.getAttribute('w:before')) styleData.spacing.before = spacing.getAttribute('w:before');
                                if (spacing.getAttribute('w:after')) styleData.spacing.after = spacing.getAttribute('w:after');
                                if (spacing.getAttribute('w:line')) styleData.spacing.line = spacing.getAttribute('w:line');
                                if (spacing.getAttribute('w:lineRule')) styleData.spacing.lineRule = spacing.getAttribute('w:lineRule');
                            }
                        }

                        // Based On (Inheritance) - simplified (just store id)
                        const basedOn = style.getElementsByTagNameNS(this.ns.w, 'basedOn')[0];
                        if (basedOn) styleData.basedOn = basedOn.getAttribute('w:val');

                        this.docStyles.set(styleId, styleData);
                    }

                    // Run Properties within Style (for Default Paragraph Font or Character Styles)
                    const rPr = style.getElementsByTagNameNS(this.ns.w, 'rPr')[0];
                    if (rPr) {
                        // Ensure styleData exists (it might be a character style, not paragraph)
                        // For now, attaching to the same styleId. 
                        // If type='character', we should handle it, but priority is Paragraph Style inheritance.
                        let styleData = this.docStyles.get(styleId) || {};
                        styleData.rPr = {};

                        const sz = rPr.getElementsByTagNameNS(this.ns.w, 'sz')[0];
                        if (sz) styleData.rPr.sz = sz.getAttribute('w:val');

                        const color = rPr.getElementsByTagNameNS(this.ns.w, 'color')[0];
                        if (color) styleData.rPr.color = color.getAttribute('w:val');

                        // Capture basedOn if we didn't already (e.g. char style)
                        const basedOn = style.getElementsByTagNameNS(this.ns.w, 'basedOn')[0];
                        if (basedOn && !styleData.basedOn) styleData.basedOn = basedOn.getAttribute('w:val');

                        this.docStyles.set(styleId, styleData);
                    }
                }
            }
        } catch (e) {
            console.warn('스타일 파싱 실패:', e);
        }
    }

    async parseNumbering(zip) {
        try {
            const numXml = await zip.file('word/numbering.xml')?.async('text');
            if (numXml) {
                const numDoc = new DOMParser().parseFromString(numXml, 'text/xml');

                // Abstract Numbering
                const abstractNums = numDoc.getElementsByTagNameNS(this.ns.w, 'abstractNum');
                for (const abs of abstractNums) {
                    const absId = abs.getAttribute('w:abstractNumId');
                    const levels = new Map();
                    const lvls = abs.getElementsByTagNameNS(this.ns.w, 'lvl');
                    for (const lvl of lvls) {
                        const ilvl = lvl.getAttribute('w:ilvl');
                        const start = parseInt(lvl.getElementsByTagNameNS(this.ns.w, 'start')[0]?.getAttribute('w:val') || '1');
                        const numFmt = lvl.getElementsByTagNameNS(this.ns.w, 'numFmt')[0]?.getAttribute('w:val') || 'decimal';
                        const lvlText = lvl.getElementsByTagNameNS(this.ns.w, 'lvlText')[0]?.getAttribute('w:val') || '';

                        // 레벨별 들여쓰기 정보 파싱
                        let lvlLeft = null;
                        let lvlHanging = null;
                        const lvlPPr = lvl.getElementsByTagNameNS(this.ns.w, 'pPr')[0];
                        if (lvlPPr) {
                            const lvlInd = lvlPPr.getElementsByTagNameNS(this.ns.w, 'ind')[0];
                            if (lvlInd) {
                                lvlLeft = lvlInd.getAttribute('w:left') || lvlInd.getAttribute('w:start');
                                lvlHanging = lvlInd.getAttribute('w:hanging');
                            }
                        }

                        levels.set(ilvl, { start, numFmt, lvlText, lvlLeft, lvlHanging });
                    }
                    this.abstractNumMap.set(absId, levels);
                }

                // Numbering Instances
                const nums = numDoc.getElementsByTagNameNS(this.ns.w, 'num');
                for (const num of nums) {
                    const numId = num.getAttribute('w:numId');
                    const absRef = num.getElementsByTagNameNS(this.ns.w, 'abstractNumId')[0]?.getAttribute('w:val');
                    if (absRef) {
                        this.numberingMap.set(numId, absRef);
                    }
                }
            }
        } catch (e) {
            console.warn('번호 매기기 파싱 실패:', e);
        }
    }

    async parseFootnotes(zip) {
        try {
            const xml = await zip.file('word/footnotes.xml')?.async('text');
            if (xml) {
                const doc = new DOMParser().parseFromString(xml, 'text/xml');
                const footnotes = doc.getElementsByTagNameNS(this.ns.w, 'footnote');
                for (const fn of footnotes) {
                    const id = fn.getAttribute('w:id');
                    this.footnoteMap.set(id, fn);
                }
            }
        } catch (e) {
            console.warn('각주 파싱 실패:', e);
        }
    }

    async parseEndnotes(zip) {
        try {
            const xml = await zip.file('word/endnotes.xml')?.async('text');
            if (xml) {
                const doc = new DOMParser().parseFromString(xml, 'text/xml');
                const endnotes = doc.getElementsByTagNameNS(this.ns.w, 'endnote');
                for (const en of endnotes) {
                    const id = en.getAttribute('w:id');
                    this.endnoteMap.set(id, en);
                }
            }
        } catch (e) {
            console.warn('미주 파싱 실패:', e);
        }
    }

    getListPrefix(numId, ilvl) {
        const absRef = this.numberingMap.get(numId);
        if (!absRef) return '';

        const levels = this.abstractNumMap.get(absRef);
        if (!levels) return '';

        const lvl = levels.get(ilvl);
        if (!lvl) return '';

        const ilvlNum = parseInt(ilvl);

        // 불릿 목록 (번호 추적 불필요)
        if (lvl.numFmt === 'bullet') {
            // 레벨에 따른 불릿 기호 변경
            const bulletSymbols = ['●', '○', '■', '◆', '▪', '▸', '‣', '⁃', '⦿'];
            if (lvl.lvlText === '' || lvl.lvlText === '') return bulletSymbols[Math.min(ilvlNum, bulletSymbols.length - 1)] + ' ';
            if (lvl.lvlText === 'o') return '○ ';
            if (lvl.lvlText === '§') return '■ ';
            return bulletSymbols[Math.min(ilvlNum, bulletSymbols.length - 1)] + ' ';
        }

        // 번호 목록 — 카운터 추적
        const counterKey = `${numId}:${ilvl}`;

        // 카운터 증가 또는 초기화
        if (this.listCounters.has(counterKey)) {
            this.listCounters.set(counterKey, this.listCounters.get(counterKey) + 1);
        } else {
            // 초기값 설정
            this.listCounters.set(counterKey, lvl.start || 1);
        }

        // 상위 레벨이 변경되면 하위 레벨 카운터 리셋
        for (let i = ilvlNum + 1; i <= 9; i++) {
            const subKey = `${numId}:${i}`;
            this.listCounters.delete(subKey);
        }

        // 다단계 번호 텍스트 생성 (%1, %2, ... 플레이스홀더 치환)
        let text = lvl.lvlText;
        for (let i = 0; i <= ilvlNum; i++) {
            const placeholder = `%${i + 1}`;
            const levelKey = `${numId}:${i}`;
            const levelVal = this.listCounters.get(levelKey) || 1;

            // 해당 레벨의 포맷 규칙 적용
            const levelDef = levels.get(String(i));
            let numStr = String(levelVal);
            if (levelDef) {
                if (levelDef.numFmt === 'lowerLetter') numStr = String.fromCharCode(96 + Math.min(levelVal, 26));
                else if (levelDef.numFmt === 'upperLetter') numStr = String.fromCharCode(64 + Math.min(levelVal, 26));
                else if (levelDef.numFmt === 'lowerRoman') numStr = this.toRoman(levelVal).toLowerCase();
                else if (levelDef.numFmt === 'upperRoman') numStr = this.toRoman(levelVal);
                else if (levelDef.numFmt === 'ganada') numStr = String.fromCharCode(0xAC00 + levelVal - 1); // 가나다
                else if (levelDef.numFmt === 'chosung') numStr = String.fromCharCode(0x3131 + levelVal - 1); // ㄱㄴㄷ
            }

            text = text.replace(placeholder, numStr);
        }

        return text + ' ';
    }

    toRoman(num) {
        // Simple Roman converter
        const roman = { M: 1000, CM: 900, D: 500, CD: 400, C: 100, XC: 90, L: 50, XL: 40, X: 10, IX: 9, V: 5, IV: 4, I: 1 };
        let str = '';
        for (const i of Object.keys(roman)) {
            const q = Math.floor(num / roman[i]);
            num -= q * roman[i];
            str += i.repeat(q);
        }
        return str;
    }

    async parseImages(zip) {
        for (const [rId, rel] of this.relationships) {
            if (rel.type?.includes('image')) {
                try {
                    const imagePath = 'word/media/' + rel.target;
                    const imageFile = zip.file(imagePath);
                    if (imageFile) {
                        const imageData = await imageFile.async('uint8array');
                        const ext = imagePath.split('.').pop().toLowerCase();
                        const manifestId = `BIN${String(this.imageCounter).padStart(4, '0')}`;

                        // 이미지 크기 추출
                        const dimensions = this.getImageDimensions(imageData, ext);

                        this.images.set(rId, {
                            data: imageData,
                            ext: ext === 'jpeg' ? 'jpg' : ext, // Standardize extension
                            manifestId: manifestId,
                            path: `BinData/${manifestId}.${ext === 'jpeg' ? 'jpg' : ext}`,
                            width: dimensions.width,
                            height: dimensions.height
                        });
                        this.imageCounter++;
                    }
                } catch (e) {
                    console.warn(`이미지 ${rId} 파싱 실패:`, e);
                }
            }
        }
        console.log(`✅ ${this.images.size}개 이미지 파싱 완료`);
    }

    getImageDimensions(data, ext) {
        let width = 0;
        let height = 0;

        try {
            if (ext === 'png') {
                // PNG Header: 89 50 4E 47 0D 0A 1A 0A
                // IHDR Chunk: Length(4) Type(4) Width(4) Height(4) ...
                // IHDR is usually the first chunk after signature
                const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
                if (view.getUint32(0) === 0x89504E47 && view.getUint32(4) === 0x0D0A1A0A) {
                    // IHDR starts at offset 8. 
                    // Length at 8 (should be 13), Type at 12 ('IHDR'), Width at 16, Height at 20
                    width = view.getUint32(16);
                    height = view.getUint32(20);
                }
            } else if (ext === 'jpg' || ext === 'jpeg') {
                const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
                let offset = 2; // Skip SOI (FF D8)
                while (offset < view.byteLength) {
                    const marker = view.getUint16(offset);
                    offset += 2;

                    if (marker === 0xFFC0 || marker === 0xFFC2) { // SOF0 or SOF2
                        // Length (2), Precision (1), Height (2), Width (2)
                        height = view.getUint16(offset + 3);
                        width = view.getUint16(offset + 5);
                        break;
                    } else {
                        const len = view.getUint16(offset);
                        offset += len;
                    }
                }
            }
        } catch (e) {
            console.warn('이미지 크기 추출 실패 (기본값 사용):', e);
        }

        // Fallback if extraction failed
        return { width: width || 100, height: height || 100 };
    }

    initDefaultStyles() {
        // BorderFills (PDF 표 23 참고)
        // BorderFill 1: 페이지 테두리
        this.borderFills.set('1', {
            id: '1',
            leftBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            rightBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            topBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            bottomBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' }
        });

        // BorderFill 2: 표 기본 테두리
        this.borderFills.set('2', {
            id: '2',
            leftBorder: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            rightBorder: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            topBorder: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            bottomBorder: { type: 'SOLID', width: '0.1 mm', color: '#000000' }
        });

        // BorderFill 3: 표 외곽 테두리
        this.borderFills.set('3', {
            id: '3',
            leftBorder: { type: 'SOLID', width: '0.5 mm', color: '#000000' },
            rightBorder: { type: 'SOLID', width: '0.5 mm', color: '#000000' },
            topBorder: { type: 'SOLID', width: '0.5 mm', color: '#000000' },
            bottomBorder: { type: 'SOLID', width: '0.5 mm', color: '#000000' }
        });

        // BorderFill 4: 표 셀 테두리
        this.borderFills.set('4', {
            id: '4',
            leftBorder: { type: 'SOLID', width: '0.1 mm', color: '#D0D0D0' },
            rightBorder: { type: 'SOLID', width: '0.1 mm', color: '#D0D0D0' },
            topBorder: { type: 'SOLID', width: '0.1 mm', color: '#D0D0D0' },
            bottomBorder: { type: 'SOLID', width: '0.1 mm', color: '#D0D0D0' }
        });

        // BorderFill 5: Box Text (박스형)
        this.borderFills.set('5', {
            id: '5',
            leftBorder: { type: 'SOLID', width: '0.12 mm', color: '#000000' },
            rightBorder: { type: 'SOLID', width: '0.12 mm', color: '#000000' },
            topBorder: { type: 'SOLID', width: '0.12 mm', color: '#000000' },
            bottomBorder: { type: 'SOLID', width: '0.12 mm', color: '#000000' }
        });

        // BorderFill 6: Inverse Text (반전)
        this.borderFills.set('6', {
            id: '6',
            leftBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            rightBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            topBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            bottomBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backColor: '#000000'
        });

        // CharShapes (PDF 표 33 참고)
        // CharShape 0: 기본
        this.charShapes.set('0', {
            id: '0', height: '1000', textColor: '#000000',
            bold: false, italic: false, underline: false
        });

        // CharShape 1: 기본 (표 내부용)
        this.charShapes.set('1', {
            id: '1', height: '1000', textColor: '#000000',
            bold: false, italic: false, underline: false
        });

        // CharShape 2: 기본 (본문용)
        this.charShapes.set('2', {
            id: '2', height: '1000', textColor: '#000000',
            bold: false, italic: false, underline: false
        });

        // CharShape 19: 굵게
        this.charShapes.set('19', {
            id: '19', height: '1000', textColor: '#000000',
            bold: true, italic: false, underline: false
        });

        // CharShape 24: 제목용 굵게 큰 글자
        this.charShapes.set('24', {
            id: '24', height: '1400', textColor: '#000000',
            bold: true, italic: false, underline: false
        });

        // CharShape 28: 코드용 작은 글자
        this.charShapes.set('28', {
            id: '28', height: '900', textColor: '#000000',
            bold: false, italic: false, underline: false
        });

        // ParaShapes (PDF 표 43 참고)
        this.paraShapes.set('0', {
            id: '0', align: 'LEFT', heading: 'NONE', level: '0'
        });
    }

    createHwpx(doc) {
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

    createVersion() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hv:HCFVersion xmlns:hv="${this.hwpxNs.hv}" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="0" buildNumber="1" os="6" xmlVersion="1.2" application="Hancom Office Hangul" appVersion="11, 20, 0, 1520"/>`;
    }

    createSettings() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ha:HWPApplicationSetting xmlns:ha="${this.hwpxNs.ha}" xmlns:config="${this.hwpxNs.config}"><ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/></ha:HWPApplicationSetting>`;
    }

    createHeader() {
        let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hh:head xmlns:ha="${this.hwpxNs.ha}" xmlns:hp="${this.hwpxNs.hp}" xmlns:hp10="${this.hwpxNs.hp10}" xmlns:hs="${this.hwpxNs.hs}" xmlns:hc="${this.hwpxNs.hc}" xmlns:hh="${this.hwpxNs.hh}" xmlns:hhs="${this.hwpxNs.hhs}" xmlns:hm="${this.hwpxNs.hm}" xmlns:hpf="${this.hwpxNs.hpf}" xmlns:dc="${this.hwpxNs.dc}" xmlns:opf="${this.hwpxNs.opf}" xmlns:ooxmlchart="${this.hwpxNs.ooxmlchart}" xmlns:epub="${this.hwpxNs.epub}" xmlns:config="${this.hwpxNs.config}" version="1.2" secCnt="1">`;
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

    // createFontFaces is defined later (dynamic version)

    createBorderFills() {
        let xml = `<hh:borderFills itemCnt="${this.borderFills.size}">`;
        for (const [id, bf] of this.borderFills) {
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

    createBinDataProperties() {
        if (this.images.size === 0) return '';
        let xml = `<hh:binDataProperties itemCnt="${this.images.size}">`;
        for (const [, imgInfo] of this.images) {
            // id는 본문에서 참조할 번호, itemIDRef는 HWPX manifest의 ID입니다.
            xml += `<hh:binData id="${imgInfo.binDataId}" itemIDRef="${imgInfo.manifestId}" extension="${imgInfo.ext}" format="${imgInfo.ext}" type="EMBEDDING" state="NOT_ACCESSED"/>`;
        }
        xml += `</hh:binDataProperties>`;
        return xml;
    }


    createTabProperties() {
        // PDF 표 36 참고 - 2개 탭 속성 정의 (기본 + 헤더/푸터용)
        let xml = `<hh:tabProperties itemCnt="2">`;
        xml += `<hh:tabPr id="0" itemCnt="0" autoTabLeft="1" autoTabRight="0"/>`;
        // 탭 속성 1: 중앙/오른쪽 정렬 탭 (헤더/푸터용)
        xml += `<hh:tabPr id="1" itemCnt="2" autoTabLeft="0" autoTabRight="0">`;
        const centerTab = Math.round((this.PAGE_WIDTH - this.MARGIN_LEFT - this.MARGIN_RIGHT) / 2);
        const rightTab = this.PAGE_WIDTH - this.MARGIN_LEFT - this.MARGIN_RIGHT;
        xml += `<hh:tabItem pos="${centerTab}" type="CENTER" leader="NONE"/>`;
        xml += `<hh:tabItem pos="${rightTab}" type="RIGHT" leader="NONE"/>`;
        xml += `</hh:tabPr>`;
        xml += `</hh:tabProperties>`;
        return xml;
    }

    /**
     * DOCX 넘버링(번호매기기) 정의를 HWPX hh:numberings로 변환
     * DOCX의 abstractNum → HWPX의 hh:numbering 매핑
     * 
     * DOCX numFmt → HWPX numFormat 매핑:
     *   decimal → DIGIT
     *   lowerRoman → ROMAN_SMALL
     *   upperRoman → ROMAN_CAPITAL
     *   lowerLetter → LATIN_SMALL
     *   upperLetter → LATIN_CAPITAL
     *   bullet → BULLET
     *   ganada → DIGIT (한글 가나다 순서)
     *   chosung → DIGIT (한글 초성 순서)
     */
    createNumberings() {
        // 넘버링 정의가 없으면 빈 요소 반환
        if (this.abstractNumMap.size === 0) {
            return `<hh:numberings itemCnt="0"/>`;
        }

        // DOCX numFmt → HWPX numFormat 변환 맵
        const numFmtMap = {
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
        const bulletSymbolMap = {
            '': '·',
            '': '·',
            'o': 'o',
            '§': '§',
            '': '·'
        };

        let xml = `<hh:numberings itemCnt="${this.abstractNumMap.size}">`;
        let numberingId = 1;

        for (const [absId, levels] of this.abstractNumMap) {
            xml += `<hh:numbering id="${numberingId}" start="${levels.get('0')?.start || 1}">`;

            // 10개 레벨 생성 (HWPX 기본 구조)
            for (let level = 1; level <= 10; level++) {
                const lvlData = levels.get(String(level - 1));

                if (lvlData) {
                    const numFormat = numFmtMap[lvlData.numFmt] || 'DIGIT';
                    const startVal = lvlData.start || 1;

                    // 들여쓰기 계산 (DOCX twip → HWPUNIT, 1twip = 2HWPUNIT)
                    const indent = lvlData.lvlLeft ? Math.round(parseInt(lvlData.lvlLeft) * 2) : (level * 2200);
                    const hanging = lvlData.lvlHanging ? Math.round(parseInt(lvlData.lvlHanging) * 2) : 3600;

                    // bullet인 경우
                    if (lvlData.numFmt === 'bullet') {
                        const symbol = bulletSymbolMap[lvlData.lvlText] || '·';
                        xml += `<hh:paraHead start="${startVal}" level="${level}" align="LEFT" useInstWidth="1" autoIndent="1"`;
                        xml += ` widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="BULLET"`;
                        xml += ` charPrIDRef="0" checkable="0">${symbol}</hh:paraHead>`;
                    } else {
                        // 숫자/문자 번호 매기기
                        // lvlText에서 HWPX 포맷 문자 생성
                        let headText = this.convertLvlTextToHwpx(lvlData.lvlText, numFormat, level);

                        xml += `<hh:paraHead start="${startVal}" level="${level}" align="LEFT" useInstWidth="1" autoIndent="1"`;
                        xml += ` widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="${numFormat}"`;
                        xml += ` charPrIDRef="0" checkable="0">${this.escapeXml(headText)}</hh:paraHead>`;
                    }
                } else {
                    // 정의되지 않은 레벨은 기본 DIGIT 사용
                    xml += `<hh:paraHead start="1" level="${level}" align="LEFT" useInstWidth="1" autoIndent="0"`;
                    xml += ` widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT"`;
                    xml += ` charPrIDRef="0" checkable="0">^N</hh:paraHead>`;
                }
            }

            xml += `</hh:numbering>`;
            numberingId++;
        }

        xml += `</hh:numberings>`;
        return xml;
    }

    /**
     * DOCX lvlText (%1, %2 등)를 HWPX paraHead 텍스트로 변환
     * HWPX에서는 ^N이 현재 레벨 번호를 나타냄
     */
    convertLvlTextToHwpx(lvlText, numFormat, level) {
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

    createStyles() {
        // DOCX 스타일을 HWPX 스타일로 의미있게 매핑
        // MS Word 호환 스타일 62개 생성 (input 원본과 동일한 구조)
        const styleDefinitions = [
            // id 0: 기본 바탕글 (Normal)
            { id: 0, type: 'PARA', name: '바탕글', engName: 'Normal', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 0 },
            // id 1: Book Title (문자 스타일)
            { id: 1, type: 'CHAR', name: 'Book Title', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 1 },
            // id 2: Caption
            { id: 2, type: 'PARA', name: 'Caption', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 2 },
            // id 3: Default Paragraph Font
            { id: 3, type: 'CHAR', name: 'Default Paragraph Font', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 3 },
            // id 4: Emphasis
            { id: 4, type: 'CHAR', name: 'Emphasis', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 4 },
            // id 5: Endnote Text Char
            { id: 5, type: 'CHAR', name: 'Endnote Text Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 5 },
            // id 6: FollowedHyperlink
            { id: 6, type: 'CHAR', name: 'FollowedHyperlink', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 6 },
            // id 7: Footer
            { id: 7, type: 'PARA', name: 'Footer', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 7 },
            // id 8: Footer Char
            { id: 8, type: 'CHAR', name: 'Footer Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 8 },
            // id 9: Footnote Text Char
            { id: 9, type: 'CHAR', name: 'Footnote Text Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 9 },
            // id 10: Header
            { id: 10, type: 'PARA', name: 'Header', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 10 },
            // id 11: Header Char
            { id: 11, type: 'CHAR', name: 'Header Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 11 },
            // id 12: Heading 1
            { id: 12, type: 'PARA', name: 'Heading 1', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 12 },
            // id 13: Heading 1 Char
            { id: 13, type: 'CHAR', name: 'Heading 1 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 13 },
            // id 14: Heading 2
            { id: 14, type: 'PARA', name: 'Heading 2', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 14 },
            // id 15: Heading 2 Char
            { id: 15, type: 'CHAR', name: 'Heading 2 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 15 },
            // id 16: Heading 3
            { id: 16, type: 'PARA', name: 'Heading 3', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 16 },
            // id 17: Heading 3 Char
            { id: 17, type: 'CHAR', name: 'Heading 3 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 17 },
            // id 18: Heading 4
            { id: 18, type: 'PARA', name: 'Heading 4', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 18 },
            // id 19: Heading 4 Char
            { id: 19, type: 'CHAR', name: 'Heading 4 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 19 },
            // id 20: Heading 5
            { id: 20, type: 'PARA', name: 'Heading 5', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 20 },
            // id 21: Heading 5 Char
            { id: 21, type: 'CHAR', name: 'Heading 5 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 21 },
            // id 22: Heading 6
            { id: 22, type: 'PARA', name: 'Heading 6', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 22 },
            // id 23: Heading 6 Char
            { id: 23, type: 'CHAR', name: 'Heading 6 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 23 },
            // id 24: Heading 7
            { id: 24, type: 'PARA', name: 'Heading 7', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 24 },
            // id 25: Heading 7 Char
            { id: 25, type: 'CHAR', name: 'Heading 7 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 25 },
            // id 26: Heading 8
            { id: 26, type: 'PARA', name: 'Heading 8', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 26 },
            // id 27: Heading 8 Char
            { id: 27, type: 'CHAR', name: 'Heading 8 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 27 },
            // id 28: Heading 9
            { id: 28, type: 'PARA', name: 'Heading 9', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 28 },
            // id 29: Heading 9 Char
            { id: 29, type: 'CHAR', name: 'Heading 9 Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 29 },
            // id 30: Hyperlink
            { id: 30, type: 'CHAR', name: 'Hyperlink', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 30 },
            // id 31: Intense Emphasis
            { id: 31, type: 'CHAR', name: 'Intense Emphasis', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 31 },
            // id 32: Intense Quote
            { id: 32, type: 'PARA', name: 'Intense Quote', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 32 },
            // id 33: Intense Quote Char
            { id: 33, type: 'CHAR', name: 'Intense Quote Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 33 },
            // id 34: Intense Reference
            { id: 34, type: 'CHAR', name: 'Intense Reference', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 34 },
            // id 35: List Paragraph
            { id: 35, type: 'PARA', name: 'List Paragraph', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 35 },
            // id 36: No Spacing
            { id: 36, type: 'PARA', name: 'No Spacing', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 36 },
            // id 37: Placeholder Text
            { id: 37, type: 'CHAR', name: 'Placeholder Text', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 37 },
            // id 38: Quote
            { id: 38, type: 'PARA', name: 'Quote', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 38 },
            // id 39: Quote Char
            { id: 39, type: 'CHAR', name: 'Quote Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 39 },
            // id 40: Strong
            { id: 40, type: 'CHAR', name: 'Strong', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 40 },
            // id 41: Subtitle
            { id: 41, type: 'PARA', name: 'Subtitle', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 41 },
            // id 42: Subtitle Char
            { id: 42, type: 'CHAR', name: 'Subtitle Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 42 },
            // id 43: Subtle Emphasis
            { id: 43, type: 'CHAR', name: 'Subtle Emphasis', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 43 },
            // id 44: Subtle Reference
            { id: 44, type: 'CHAR', name: 'Subtle Reference', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 44 },
            // id 45: TOC Heading
            { id: 45, type: 'PARA', name: 'TOC Heading', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 45 },
            // id 46: Title
            { id: 46, type: 'PARA', name: 'Title', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 46 },
            // id 47: Title Char
            { id: 47, type: 'CHAR', name: 'Title Char', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 47 },
            // id 48: endnote reference
            { id: 48, type: 'CHAR', name: 'endnote reference', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 48 },
            // id 49: endnote text
            { id: 49, type: 'PARA', name: 'endnote text', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 49 },
            // id 50: footnote reference
            { id: 50, type: 'CHAR', name: 'footnote reference', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 50 },
            // id 51: footnote text
            { id: 51, type: 'PARA', name: 'footnote text', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 51 },
            // id 52: table of figures
            { id: 52, type: 'PARA', name: 'table of figures', engName: '', paraPrIDRef: 0, charPrIDRef: 0, nextStyleIDRef: 52 },
            // id 53~61: toc 1~9
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
        this.docxStyleToHwpxId = {
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
            xml += `<hh:style id="${s.id}" type="${s.type}" name="${this.escapeXml(s.name)}" engName="${this.escapeXml(s.engName)}" paraPrIDRef="${s.paraPrIDRef}" charPrIDRef="${s.charPrIDRef}" nextStyleIDRef="${s.nextStyleIDRef}" langID="1042" lockForm="0"/>`;
        }
        xml += `</hh:styles>`;
        return xml;
    }

    createSection(doc) {
        let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec xmlns:ha="${this.hwpxNs.ha}" xmlns:hp="${this.hwpxNs.hp}" xmlns:hp10="${this.hwpxNs.hp10}" xmlns:hs="${this.hwpxNs.hs}" xmlns:hc="${this.hwpxNs.hc}" xmlns:hh="${this.hwpxNs.hh}" xmlns:hhs="${this.hwpxNs.hhs}" xmlns:hm="${this.hwpxNs.hm}" xmlns:hpf="${this.hwpxNs.hpf}" xmlns:dc="${this.hwpxNs.dc}" xmlns:opf="${this.hwpxNs.opf}" xmlns:ooxmlchart="${this.hwpxNs.ooxmlchart}" xmlns:epub="${this.hwpxNs.epub}" xmlns:config="${this.hwpxNs.config}">`;

        const body = doc.getElementsByTagNameNS(this.ns.w, 'body')[0];
        if (body) {
            const paraCount = body.getElementsByTagNameNS(this.ns.w, 'p').length;
            console.log(`Converting ${paraCount} paragraphs...`);

            // 본문의 모든 자식 요소를 배열로 변환하여 마지막 요소 판별
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
                    // SDT (Structured Document Tag) — 목차(TOC), 콘텐츠 컨트롤 등
                    // 내부 sdtContent에서 문단/테이블 추출
                    const sdtContent = child.getElementsByTagNameNS(this.ns.w, 'sdtContent')[0];
                    if (sdtContent) {
                        const sdtChildren = Array.from(sdtContent.children).filter(
                            c => c.localName === 'p' || c.localName === 'tbl'
                        );
                        const totalSdtChildren = sdtChildren.length;
                        sdtChildren.forEach((sdtChild, sdtIndex) => {
                            // SDT 내부의 마지막 요소이면서 전체 본문의 마지막 SDT인 경우
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
            // 빈 문서
            const paraId = this.idCounter++;
            xml += `<hp:p id="${paraId}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="2">${this.createSecPr()}<hp:t></hp:t></hp:run><hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1265" textheight="1265" baseline="1032" spacing="188" horzpos="0" horzsize="${this.TEXT_WIDTH}" flags="393216"/></hp:linesegarray></hp:p>`;
        }
        xml += `</hs:sec>`;
        return xml;
    }

    createSecPr() {
        // PDF 표 129 참고 - 구역 정의
        // PDF 표 131 참고 - 용지 설정
        return `<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="${this.COL_GAP}" tabStop="8000" outlineShapeIDRef="1" memoShapeIDRef="1" textVerticalWidthHead="0" masterPageCnt="0"><hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/><hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/><hp:pagePr landscape="WIDELY" width="${this.PAGE_WIDTH}" height="${this.PAGE_HEIGHT}" gutterType="LEFT_ONLY"><hp:margin header="${this.HEADER_MARGIN}" footer="${this.FOOTER_MARGIN}" gutter="${this.GUTTER_MARGIN}" left="${this.MARGIN_LEFT}" right="${this.MARGIN_RIGHT}" top="${this.MARGIN_TOP}" bottom="${this.MARGIN_BOTTOM}"/></hp:pagePr><hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/><hp:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/><hp:noteSpacing betweenNotes="283" belowLine="0" aboveLine="1000"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr><hp:endNotePr><hp:autoNumFormat type="ROMAN_SMALL" userChar="" prefixChar="" suffixChar="" supscript="1"/><hp:noteLine length="-1" type="SOLID" width="0.25 mm" color="#000000"/><hp:noteSpacing betweenNotes="0" belowLine="0" aboveLine="1000"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr><hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill><hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417" right="1417" top="1417" bottom="1417"/></hp:pageBorderFill></hp:secPr><hp:ctrl><hp:colPr id="" type="${this.COL_TYPE}" layout="LEFT" colCount="${this.COL_COUNT}" sameSz="1" sameGap="0"/></hp:ctrl>`;
    }

    convertParagraph(para, isLastElement = false) {
        const paraId = this.idCounter++;

        // 문단 모양(ParaShape) 생성 및 ID 획득
        const paraPrId = this.getParaShapeId(para, isLastElement);

        // 문단 스타일 ID
        let styleId = '0';
        let paraStyleId = null;

        const pPr = para.getElementsByTagNameNS(this.ns.w, 'pPr')[0];
        if (pPr) {
            const pStyle = pPr.getElementsByTagNameNS(this.ns.w, 'pStyle')[0];
            if (pStyle) {
                paraStyleId = pStyle.getAttribute('w:val');
                // styleId = paraStyleId; // HWPX style mapping not fully implemented yet, keeping '0' for safety or ID mapping needs to be done.
                // For now, we use styleId='0' for the XML attribute but use paraStyleId for property lookup.
            }
        }

        // 줄 간격 및 높이 계산을 위한 대표 폰트 크기
        // 문단 내 첫 번째 런의 폰트 크기나 기본값 사용
        const baseFontSize = this.getParagraphFontSize(para);
        const fontSizeInHWPUNIT = Math.round(baseFontSize * this.HWPUNIT_PER_INCH / 72);

        // 텍스트 높이 및 줄 간격 계산
        // paraShape에서 실제 줄간격 비율 사용
        const paraShape = this.paraShapes.get(paraPrId);
        const lineSpacingPercent = (paraShape && paraShape.lineSpacingType === 'PERCENT') ? paraShape.lineSpacingVal : 115;
        const textheight = fontSizeInHWPUNIT;
        const baseline = Math.round(textheight * 0.85); // 베이스라인 조정
        const spacing = Math.round(textheight * (lineSpacingPercent - 100) / 100); // 실제 줄간격 비율 적용
        const vertsize = textheight + spacing; // 전체 높이

        // Page Break & CharShape ID Calculation
        let pageBreak = '0';
        let charPrId = '0'; // Default charPrId for paragraph markers/prefixes

        if (pPr) {
            // Page Break
            const pageBreakBefore = pPr.getElementsByTagNameNS(this.ns.w, 'pageBreakBefore')[0];
            if (pageBreakBefore && pageBreakBefore.getAttribute('w:val') !== 'false') {
                pageBreak = '1';
            }

            // CharShape properteis from paragraph style
            const rPr = pPr.getElementsByTagNameNS(this.ns.w, 'rPr')[0];
            charPrId = this.getCharShapeId(rPr, baseFontSize, false, paraStyleId);
        } else {
            // No pPr, use default with base font size
            charPrId = this.getCharShapeId(null, baseFontSize, false, paraStyleId);
        }

        // DOCX 스타일 ID → HWPX 스타일 ID 매핑
        let hwpxStyleId = '0';
        if (paraStyleId && this.docxStyleToHwpxId && this.docxStyleToHwpxId[paraStyleId] !== undefined) {
            hwpxStyleId = String(this.docxStyleToHwpxId[paraStyleId]);
        }

        let xml = `<hp:p id="${paraId}" paraPrIDRef="${paraPrId}" styleIDRef="${hwpxStyleId}" pageBreak="${pageBreak}" columnBreak="0" merged="0">`;

        // 첫 문단에 secPr 추가
        if (this.isFirstParagraph) {
            // 빈 런을 만들어 secPr을 포함시킴
            xml += `<hp:run charPrIDRef="0">${this.createSecPr()}</hp:run>`;
            this.isFirstParagraph = false;
        }

        // 목록(Bullets/Numbering) 처리
        // HWPX에서는 paraPr의 heading type="OUTLINE" + idRef로 자동 표시되므로
        // 텍스트 접두어를 직접 삽입하지 않음 (getParaShapeId에서 heading 설정됨)
        const numPr = para.getElementsByTagNameNS(this.ns.w, 'numPr')[0];
        if (numPr) {
            const numId = numPr.getElementsByTagNameNS(this.ns.w, 'numId')[0]?.getAttribute('w:val');
            const ilvl = numPr.getElementsByTagNameNS(this.ns.w, 'ilvl')[0]?.getAttribute('w:val');
            // 카운터 추적만 유지 (getListPrefix 내부에서 카운터 업데이트)
            if (numId && ilvl) {
                this.getListPrefix(numId, ilvl);
            }
        }

        // Run(텍스트, 이미지 등) 변환
        // DOCX의 Run들을 순회하며 HWPX Run으로 변환
        const runs = para.childNodes;
        for (const node of runs) {
            if (node.nodeName === 'w:r') {
                xml += this.convertRun(node, baseFontSize, paraStyleId);
            } else if (node.nodeName === 'w:hyperlink') {
                // 하이퍼링크 내부 Run 처리
                for (const child of node.childNodes) {
                    if (child.nodeName === 'w:r') {
                        // TODO: 하이퍼링크 속성 적용 필요 (파란색, 밑줄 등)
                        xml += this.convertRun(child, baseFontSize, paraStyleId);
                    }
                }
            }
        }

        // 그림 개체 처리 (`w:drawing`이 `w:r` 내부에 있지 않고 `w:p` 직계 자식인 경우 - 드물지만 대비)
        const drawings = para.getElementsByTagNameNS(this.ns.w, 'drawing');
        for (const drawing of drawings) {
            // drawing이 run 내부에 있는 경우는 convertRun에서 처리됨. 
            // 여기서는 중복 처리를 방지해야 함. 
            // DOM 순회 시 w:r 내의 drawing을 처리하므로, 여기서는 별도 처리 안 함.
        }

        // 문단 전후 간격 반영
        const actualPrevSpacing = (paraShape && paraShape.prevSpacing) ? paraShape.prevSpacing : 0;
        const actualNextSpacing = (paraShape && paraShape.nextSpacing) ? paraShape.nextSpacing : 0;

        // 문단 앞 간격을 현재 위치에 반영
        this.currentVertPos += actualPrevSpacing;

        // LineSeg (줄 정보)
        // 실제로는 텍스트 길이에 따라 여러 줄로 나뉘어야 하지만, 
        // 여기서는 문단 전체를 하나의 줄로 단순화하여 처리 (복잡한 줄바꿈 계산 로직 생략)
        xml += `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${this.currentVertPos}" vertsize="${vertsize}" textheight="${textheight}" baseline="${baseline}" spacing="${spacing}" horzpos="0" horzsize="${this.TEXT_WIDTH}" flags="393216"/></hp:linesegarray></hp:p>`;

        // 다음 문단 위치 누적 (텍스트 높이 + 줄간격 + 문단 뒤 간격)
        this.currentVertPos += vertsize + actualNextSpacing;

        return xml;
    }

    convertRun(run, baseFontSize, paraStyleId = null) {
        // Run 속성 파싱하여 CharShape ID 획득
        const rPr = run.getElementsByTagNameNS(this.ns.w, 'rPr')[0];

        // 부모가 hyperlink인 경우 스타일 강제 적용
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
                // 텍스트 런 생성
                xml += `<hp:run charPrIDRef="${charPrId}">`;
                xml += textBuffer;
                xml += `</hp:run>`;
                textBuffer = '';
                hasContentInRun = false;
            }
        };

        // DOCX Run의 자식 노드들을 순서대로 처리 (preserve order)
        // w:t (text), w:tab (tab), w:br (break), w:drawing (image), w:footnoteReference, w:endnoteReference
        for (const child of run.childNodes) {
            if (child.nodeName === 'w:t') {
                textBuffer += `<hp:t>${this.escapeXml(child.textContent)}</hp:t>`;
                hasContentInRun = true;
            } else if (child.nodeName === 'w:tab') {
                textBuffer += `<hp:t>\t</hp:t>`;
                hasContentInRun = true;
            } else if (child.nodeName === 'w:br') {
                if (child.getAttribute('w:type') === 'page') {
                    // Page break handling if needed
                } else {
                    textBuffer += `<hp:t>\n</hp:t>`;
                    hasContentInRun = true;
                }
            } else if (child.nodeName === 'w:footnoteReference') {
                const id = child.getAttribute('w:id');
                textBuffer += this.convertFootnote(id, 'FOOTNOTE');
                hasContentInRun = true;
            } else if (child.nodeName === 'w:endnoteReference') {
                const id = child.getAttribute('w:id');
                textBuffer += this.convertFootnote(id, 'ENDNOTE');
                hasContentInRun = true;
            } else if (child.nodeName === 'w:drawing') {
                // 이미지는 별도의 hp:run으로 처리되어야 하므로 현재 텍스트 버퍼를 비운다.
                flushTextBuffer();

                // 이미지는 convertDrawing이 <hp:run>...</hp:run>을 반환함
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

        // 남은 텍스트 버퍼 처리
        flushTextBuffer();

        // 만약 아무 내용도 없었다면 (빈 Run)
        // 하지만 이미지가 있었다면 xml에 이미 추가되었음.
        // 텍스트도 없고 이미지도 없는 경우만 스킵

        return xml;
    }

    convertFootnote(id, type) {
        const map = type === 'FOOTNOTE' ? this.footnoteMap : this.endnoteMap;
        const node = map.get(id);
        if (!node) return '';

        // 순차 번호 할당
        const isFootnote = type === 'FOOTNOTE';
        const number = isFootnote ? ++this.footnoteNumber : ++this.endnoteNumber;
        const instId = Math.floor(Math.random() * 10000000);

        // 각주/미주 내용 변환
        const paras = node.getElementsByTagNameNS(this.ns.w, 'p');
        let contentXml = '';

        // 첫 문단에 autoNum 삽입
        let isFirstPara = true;
        for (const para of paras) {
            if (isFirstPara) {
                // 각주 내부 첫 문단에 hp:autoNum 컨트롤 삽입
                const paraContent = this.convertParagraph(para);
                // 첫 hp:run 안에 hp:ctrl > hp:autoNum 삽입
                const autoNumType = isFootnote ? 'FOOTNOTE' : 'ENDNOTE';
                const autoNumXml = `<hp:ctrl><hp:autoNum num="${number}" numType="${autoNumType}"><hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar="" supscript="1"/></hp:autoNum></hp:ctrl>`;
                // hp:run 내부 첫 hp:t 앞에 autoNum 삽입
                contentXml += paraContent.replace(/(<hp:run[^>]*>)/, `$1${autoNumXml}`);
                isFirstPara = false;
            } else {
                contentXml += this.convertParagraph(para);
            }
        }

        // HWPX 참조 구조: <hp:ctrl><hp:footNote number="1" instId="...">...</hp:footNote></hp:ctrl>
        const tagName = isFootnote ? 'hp:footNote' : 'hp:endNote';
        const flagAttr = isFootnote ? '' : ' flag="3"';
        return `<hp:ctrl><${tagName} number="${number}" instId="${instId}"${flagAttr}><hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">${contentXml}</hp:subList></${tagName}></hp:ctrl>`;
    }

    getParaShapeId(para, forceZeroNextSpacing = false) {
        if (!para || typeof para.getElementsByTagNameNS !== 'function') return '1'; // Default ID

        // 문단 속성 추출
        const pPr = para.getElementsByTagNameNS(this.ns.w, 'pPr')[0];
        let align = 'LEFT';
        let heading = 'NONE';
        let level = '0';
        let headingIdRef = '0'; // numbering 참조 ID

        let leftMargin = 0;
        let rightMargin = 0;
        let indent = 0;
        let prevSpacing = 0;
        let nextSpacing = 0;

        let lineSpacingType = 'PERCENT';
        let lineSpacingVal = 115;

        // 줄바꿈 제어 속성
        let keepWithNext = '0';
        let keepLines = '0';
        let pageBreakBefore = '0';

        let borderFillIDRef = '1'; // Default: No border

        let hasDirectInd = false;
        let hasDirectSpacing = false;
        let paraStyleId = null;

        if (pPr) {
            // 정렬 (jc)
            let jcVal = null;
            const jc = pPr.getElementsByTagNameNS(this.ns.w, 'jc')[0];
            if (jc) {
                jcVal = jc.getAttribute('w:val');
            } else {
                // Style Lookup (Recursive)
                const pStyle = pPr.getElementsByTagNameNS(this.ns.w, 'pStyle')[0];
                if (pStyle) {
                    paraStyleId = pStyle.getAttribute('w:val');
                    let currentStyleId = paraStyleId;
                    while (currentStyleId) {
                        const style = this.docStyles.get(currentStyleId);
                        if (!style) break;

                        if (style.align) {
                            jcVal = style.align;
                            break;
                        }
                        currentStyleId = style.basedOn;
                    }
                }
            }

            if (jcVal === 'center') align = 'CENTER';
            else if (jcVal === 'right') align = 'RIGHT';
            else if (jcVal === 'both') align = 'JUSTIFY';
            else if (jcVal === 'distribute') align = 'DISTRIBUTE';

            // 들여쓰기 (ind) - 직접 속성
            const ind = pPr.getElementsByTagNameNS(this.ns.w, 'ind')[0];
            if (ind) {
                hasDirectInd = true;
                const left = ind.getAttribute('w:left') || ind.getAttribute('w:start');
                const right = ind.getAttribute('w:right') || ind.getAttribute('w:end');
                const firstLine = ind.getAttribute('w:firstLine');
                const hanging = ind.getAttribute('w:hanging');

                if (left) leftMargin = parseInt(left) * 5;
                if (right) rightMargin = parseInt(right) * 5;

                if (firstLine) {
                    indent = parseInt(firstLine) * 5;
                } else if (hanging) {
                    indent = -parseInt(hanging) * 5;
                }
            }

            // 목록 처리 (numPr) - heading/idRef를 numbering에 연결
            const numPr = pPr.getElementsByTagNameNS(this.ns.w, 'numPr')[0];
            if (numPr) {
                const numId = numPr.getElementsByTagNameNS(this.ns.w, 'numId')[0]?.getAttribute('w:val');
                const ilvl = numPr.getElementsByTagNameNS(this.ns.w, 'ilvl')[0]?.getAttribute('w:val') || '0';

                if (numId) {
                    const absRef = this.numberingMap.get(numId);
                    if (absRef) {
                        // HWPX numbering ID 계산 (abstractNumMap 순번)
                        let numberingIdx = 1;
                        for (const [key] of this.abstractNumMap) {
                            if (key === absRef) break;
                            numberingIdx++;
                        }
                        // heading을 OUTLINE으로 설정하고, idRef를 numbering ID로 연결
                        heading = 'OUTLINE';
                        headingIdRef = String(numberingIdx);
                        level = String(parseInt(ilvl) + 1); // HWPX level은 1부터 시작

                        // 들여쓰기 처리
                        if (!hasDirectInd) {
                            const levels = this.abstractNumMap.get(absRef);
                            if (levels) {
                                const lvl = levels.get(ilvl);
                                if (lvl && lvl.lvlLeft) {
                                    leftMargin = parseInt(lvl.lvlLeft) * 5;
                                    if (lvl.lvlHanging) {
                                        indent = -parseInt(lvl.lvlHanging) * 5;
                                    }
                                } else {
                                    const ilvlNum = parseInt(ilvl);
                                    leftMargin = (ilvlNum + 1) * 800 * 5;
                                    indent = -400 * 5;
                                }
                            }
                        }
                    }
                }
            }

            // 줄 간격 (spacing) - 직접 속성
            const spacing = pPr.getElementsByTagNameNS(this.ns.w, 'spacing')[0];
            if (spacing) {
                hasDirectSpacing = true;
                const before = spacing.getAttribute('w:before');
                const after = spacing.getAttribute('w:after');
                if (before) prevSpacing = parseInt(before) * 5;
                // 2/18 강제 줄간격 0으로 처리
                if (after) nextSpacing = 0;

                const line = spacing.getAttribute('w:line');
                const lineRule = spacing.getAttribute('w:lineRule');

                if (line) {
                    if (lineRule === 'auto' || !lineRule) {
                        // 240ths of a line -> Percent
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
        } // End if (pPr)

        // 2. Style Inheritance (if not direct)
        if ((!hasDirectInd || !hasDirectSpacing) && paraStyleId) {
            let currentStyleId = paraStyleId;
            while (currentStyleId) {
                const style = this.docStyles.get(currentStyleId);
                if (!style) break;

                // 들여쓰기 상속
                if (!hasDirectInd && style.ind) {
                    if (style.ind.left) leftMargin = parseInt(style.ind.left) * 5;
                    if (style.ind.right) rightMargin = parseInt(style.ind.right) * 5;
                    if (style.ind.firstLine) indent = parseInt(style.ind.firstLine) * 5;
                    if (style.ind.hanging) indent = -parseInt(style.ind.hanging) * 5;
                    hasDirectInd = true;
                }

                // 줄간격 상속
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
            const outlineLvl = pPr.getElementsByTagNameNS(this.ns.w, 'outlineLvl')[0];
            if (outlineLvl) {
                heading = 'OUTLINE';
                level = outlineLvl.getAttribute('w:val');
            }

            // keepWithNext (다음 문단과 함께 유지)
            const kwn = pPr.getElementsByTagNameNS(this.ns.w, 'keepNext')[0];
            if (kwn && kwn.getAttribute('w:val') !== 'false' && kwn.getAttribute('w:val') !== '0') {
                keepWithNext = '1';
            }

            // keepLines (단락 유지)
            const kl = pPr.getElementsByTagNameNS(this.ns.w, 'keepLines')[0];
            if (kl && kl.getAttribute('w:val') !== 'false' && kl.getAttribute('w:val') !== '0') {
                keepLines = '1';
            }

            // pageBreakBefore (페이지 나누기)
            const pbb = pPr.getElementsByTagNameNS(this.ns.w, 'pageBreakBefore')[0];
            if (pbb && pbb.getAttribute('w:val') !== 'false' && pbb.getAttribute('w:val') !== '0') {
                pageBreakBefore = '1';
            }

            // 테두리/배경 (pBdr / shd)
            const pBdr = pPr.getElementsByTagNameNS(this.ns.w, 'pBdr')[0];
            const shd = pPr.getElementsByTagNameNS(this.ns.w, 'shd')[0];

            if (pBdr || shd) {
                const parseBorder = (node) => {
                    if (!node) return { type: 'NONE', width: '0.1 mm', color: '#000000' };
                    const val = node.getAttribute('w:val');
                    if (val === 'none' || val === 'nil') return { type: 'NONE', width: '0.1 mm', color: '#000000' };

                    const sz = node.getAttribute('w:sz');
                    let width = '0.1 mm';
                    if (sz) {
                        const pt = parseInt(sz) / 8;
                        width = (pt * 0.3528).toFixed(2) + ' mm';
                    }

                    const colorAttr = node.getAttribute('w:color');
                    let color = '#000000';
                    if (colorAttr && colorAttr !== 'auto') color = `#${colorAttr}`;

                    let type = 'SOLID';
                    if (val === 'double') type = 'DOUBLE';
                    else if (val === 'dotted') type = 'DOT';
                    else if (val === 'dashed') type = 'DASH';

                    return { type, width, color };
                };

                let left = { type: 'NONE', width: '0.1 mm', color: '#000000' };
                let right = { type: 'NONE', width: '0.1 mm', color: '#000000' };
                let top = { type: 'NONE', width: '0.1 mm', color: '#000000' };
                let bottom = { type: 'NONE', width: '0.1 mm', color: '#000000' };

                if (pBdr) {
                    left = parseBorder(pBdr.getElementsByTagNameNS(this.ns.w, 'left')[0]);
                    right = parseBorder(pBdr.getElementsByTagNameNS(this.ns.w, 'right')[0]);
                    top = parseBorder(pBdr.getElementsByTagNameNS(this.ns.w, 'top')[0]);
                    bottom = parseBorder(pBdr.getElementsByTagNameNS(this.ns.w, 'bottom')[0]);
                }

                let backColor = '#FFFFFF';
                if (shd) {
                    const fill = shd.getAttribute('w:fill');
                    if (fill && fill !== 'auto') backColor = `#${fill}`;
                }

                // Generate Key
                const bfKey = `${left.type}_${left.width}_${left.color}|${right.type}_${right.width}_${right.color}|${top.type}_${top.width}_${top.color}|${bottom.type}_${bottom.width}_${bottom.color}|${backColor}`;

                let foundId = null;
                for (const [id, bf] of this.borderFills) {
                    const bfKeyCheck = `${bf.leftBorder.type}_${bf.leftBorder.width}_${bf.leftBorder.color}|${bf.rightBorder.type}_${bf.rightBorder.width}_${bf.rightBorder.color}|${bf.topBorder.type}_${bf.topBorder.width}_${bf.topBorder.color}|${bf.bottomBorder.type}_${bf.bottomBorder.width}_${bf.bottomBorder.color}|${bf.backColor || '#FFFFFF'}`;
                    if (bfKey === bfKeyCheck) {
                        foundId = id;
                        break;
                    }
                }

                if (foundId) {
                    borderFillIDRef = foundId;
                } else {
                    const newBfId = String(this.borderFills.size + 1);
                    this.borderFills.set(newBfId, {
                        id: newBfId,
                        leftBorder: left,
                        rightBorder: right,
                        topBorder: top,
                        bottomBorder: bottom,
                        backColor: backColor
                    });
                    borderFillIDRef = newBfId;
                }
            }
        }


        if (forceZeroNextSpacing) {
            nextSpacing = 0;
        }

        const key = `${align}_${leftMargin}_${rightMargin}_${indent}_${prevSpacing}_${nextSpacing}_${lineSpacingType}_${lineSpacingVal}_${heading}_${level}_${headingIdRef}_${borderFillIDRef}_${keepWithNext}_${keepLines}_${pageBreakBefore}`;

        for (const [id, ps] of this.paraShapes) {
            if (ps._key === key) return id;
        }

        const newId = String(this.paraShapes.size);
        this.paraShapes.set(newId, {
            id: newId,
            align,
            leftMargin,
            rightMargin,
            indent,
            prevSpacing,
            nextSpacing,
            lineSpacingType,
            lineSpacingVal,
            heading,
            level,
            headingIdRef,
            borderFillIDRef,
            keepWithNext,
            keepLines,
            pageBreakBefore,
            _key: key
        });
        return newId;
    }

    createParaProperties() {
        let xml = `<hh:paraProperties itemCnt="${this.paraShapes.size}">`;
        for (const [id, ps] of this.paraShapes) {
            // paraPr 속성 - 참조 형식에 맞춤
            xml += `<hh:paraPr id="${id}" tabPrIDRef="${ps.tabPrIDRef || '1'}" condense="0" fontLineHeight="0" snapToGrid="1" suppressLineNumbers="0" checked="0">`;

            // 정렬 (자식 엘리먼트)
            xml += `<hh:align horizontal="${ps.align}" vertical="BASELINE"/>`;

            // 개요 (자식 엘리먼트)
            xml += `<hh:heading type="${ps.heading}" idRef="${ps.headingIdRef || '0'}" level="${ps.level}"/>`;

            // 줄바꿈 설정
            const keepWithNext = ps.keepWithNext || '0';
            const keepLines = ps.keepLines || '0';
            const pageBreakBefore = ps.pageBreakBefore || '0';
            xml += `<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" widowOrphan="0" keepWithNext="${keepWithNext}" keepLines="${keepLines}" pageBreakBefore="${pageBreakBefore}" lineWrap="BREAK"/>`;

            // 자동 간격
            xml += `<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>`;

            // 여백
            xml += `<hh:margin>`;
            xml += `<hc:intent value="${ps.indent || 0}" unit="HWPUNIT"/>`;
            xml += `<hc:left value="${ps.leftMargin || 0}" unit="HWPUNIT"/>`;
            xml += `<hc:right value="${ps.rightMargin || 0}" unit="HWPUNIT"/>`;
            xml += `<hc:prev value="${ps.prevSpacing || 0}" unit="HWPUNIT"/>`;
            xml += `<hc:next value="${ps.nextSpacing || 0}" unit="HWPUNIT"/>`;
            xml += `</hh:margin>`;

            // 줄간격
            xml += `<hh:lineSpacing type="${ps.lineSpacingType}" value="${ps.lineSpacingVal}" unit="HWPUNIT"/>`;

            // 테두리 (참조 기본 오프셋: 400/400/100/100)
            xml += `<hh:border borderFillIDRef="${ps.borderFillIDRef}" offsetLeft="400" offsetRight="400" offsetTop="100" offsetBottom="100" connect="0" ignoreMargin="0"/>`;

            xml += `</hh:paraPr>`;
        }
        xml += `</hh:paraProperties>`;
        return xml;
    }
    createCharProperties() {
        let xml = `<hh:charProperties itemCnt="${this.charShapes.size}">`;
        for (const [id, cs] of this.charShapes) {
            // 기본 속성
            xml += `<hh:charPr id="${id}" height="${cs.height}" textColor="${cs.textColor}" shadeColor="${cs.shadeColor}" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="${cs.borderFillIDRef}">`;


            // FontRef (언어별 개별 폰트 ID 사용)
            const fIds = cs.fontIds || { hangulId: cs.fontId || 0, latinId: cs.fontId || 0, hanjaId: cs.fontId || 0, japaneseId: cs.fontId || 0, otherId: cs.fontId || 0, symbolId: cs.fontId || 0, userId: cs.fontId || 0 };
            xml += `<hh:fontRef hangul="${fIds.hangulId}" latin="${fIds.latinId}" hanja="${fIds.hanjaId}" japanese="${fIds.japaneseId}" other="${fIds.otherId}" symbol="${fIds.symbolId}" user="${fIds.userId}"/>`;

            // Ratio
            xml += `<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>`;

            // Spacing
            xml += `<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>`;

            // RelSz
            xml += `<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>`;

            // Offset
            xml += `<hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>`;

            // Child Elements for Formatting
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

    /**
     * 형광색/배경색용 borderFill을 찾거나 새로 생성
     * 기존에 동일한 faceColor를 가진 borderFill이 있으면 해당 ID 반환,
     * 없으면 테두리 없이 배경색만 설정된 borderFill을 동적으로 생성
     */
    getOrCreateHighlightBorderFill(highlightColor) {
        // 기존 borderFill에서 동일한 배경색을 가진 항목 검색
        for (const [id, bf] of this.borderFills) {
            if (bf.backColor === highlightColor &&
                bf.leftBorder && bf.leftBorder.type === 'NONE' &&
                bf.rightBorder && bf.rightBorder.type === 'NONE' &&
                bf.topBorder && bf.topBorder.type === 'NONE' &&
                bf.bottomBorder && bf.bottomBorder.type === 'NONE') {
                return id;
            }
        }

        // 새로운 형광색 borderFill 생성 (테두리 없이 배경색만)
        const newBfId = String(this.borderFills.size + 1);
        this.borderFills.set(newBfId, {
            id: newBfId,
            leftBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            rightBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            topBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            bottomBorder: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backColor: highlightColor
        });
        return newBfId;
    }

    getCharShapeId(rPr, baseFontSize, isHyperlink = false, paraStyleId = null) {
        let bold = false;
        let italic = false;
        let underline = isHyperlink;
        let strike = false;
        let supscript = false;
        let subscript = false;

        // Color defaults
        let color = isHyperlink ? '#0000FF' : '#000000';
        let shadeColor = 'none'; // Highlight
        let borderFillIDRef = '1'; // Default (No border/fill)

        let fontSize = baseFontSize ? Math.round(baseFontSize * 100) : 1000; // 기본 10pt (1000 HWPX 단위, pt × 100)
        let fontFace = '함초롬바탕';
        let directRFonts = null; // rFonts 엘리먼트 참조 (언어별 폰트 등록용)

        // Direct Formatting
        if (rPr) {
            // ... existing rPr parsing ...
            // We need to keep the existing parsing logic but merge it with style lookup for size
        }

        // 1. Parse Direct Formatting first (to override)
        let hasDirectSize = false;
        if (rPr) {
            // ... (Copy existing parsing logic here for bold, italic, etc)
            // Bold
            if (rPr.getElementsByTagNameNS(this.ns.w, 'b').length > 0 && rPr.getElementsByTagNameNS(this.ns.w, 'b')[0].getAttribute('w:val') !== '0') bold = true;

            // Italic
            if (rPr.getElementsByTagNameNS(this.ns.w, 'i').length > 0 && rPr.getElementsByTagNameNS(this.ns.w, 'i')[0].getAttribute('w:val') !== '0') italic = true;

            // Underline
            if (rPr.getElementsByTagNameNS(this.ns.w, 'u').length > 0 && rPr.getElementsByTagNameNS(this.ns.w, 'u')[0].getAttribute('w:val') !== 'none') underline = true;

            // Strike
            if (rPr.getElementsByTagNameNS(this.ns.w, 'strike').length > 0 && rPr.getElementsByTagNameNS(this.ns.w, 'strike')[0].getAttribute('w:val') !== '0') strike = true;
            if (rPr.getElementsByTagNameNS(this.ns.w, 'dstrike').length > 0 && rPr.getElementsByTagNameNS(this.ns.w, 'dstrike')[0].getAttribute('w:val') !== '0') strike = true;

            // VertAlign
            const vertAlign = rPr.getElementsByTagNameNS(this.ns.w, 'vertAlign')[0];
            if (vertAlign) {
                const val = vertAlign.getAttribute('w:val');
                if (val === 'superscript') supscript = true;
                else if (val === 'subscript') subscript = true;
            }

            // Text Color
            const colorElem = rPr.getElementsByTagNameNS(this.ns.w, 'color')[0];
            if (colorElem) {
                const val = colorElem.getAttribute('w:val');
                if (val && val !== 'auto') {
                    color = `#${val}`;
                }
            }

            // Size
            const sz = rPr.getElementsByTagNameNS(this.ns.w, 'sz')[0];
            if (sz) {
                const val = parseInt(sz.getAttribute('w:val'));
                if (!isNaN(val)) {
                    fontSize = val * 50;
                    hasDirectSize = true;
                }
            }

            // ... (Highlight, Shading, Border, Font logic - keep as is)
            // Highlight (형광색)
            const highlight = rPr.getElementsByTagNameNS(this.ns.w, 'highlight')[0];
            if (highlight) {
                const val = highlight.getAttribute('w:val');
                const highlightMap = {
                    'black': '#000000', 'blue': '#0000FF', 'cyan': '#00FFFF', 'green': '#008000',
                    'magenta': '#FF00FF', 'red': '#FF0000', 'yellow': '#FFFF00', 'white': '#FFFFFF',
                    'darkBlue': '#00008B', 'darkCyan': '#008B8B', 'darkGreen': '#006400',
                    'darkMagenta': '#8B008B', 'darkRed': '#8B0000', 'darkYellow': '#808000',
                    'darkGray': '#A9A9A9', 'lightGray': '#D3D3D3', 'none': 'none'
                };
                if (highlightMap[val] && highlightMap[val] !== 'none') {
                    shadeColor = highlightMap[val];
                    // 형광색용 borderFill 생성 (input 원본과 동일한 방식)
                    borderFillIDRef = this.getOrCreateHighlightBorderFill(shadeColor);
                }
            }

            // Shading (배경색) via w:shd
            const shd = rPr.getElementsByTagNameNS(this.ns.w, 'shd')[0];
            if (shd) {
                const fill = shd.getAttribute('w:fill');
                if (fill && fill !== 'auto') {
                    // 검정 배경 → 반전 텍스트
                    if (fill === '000000' || fill === 'black') {
                        borderFillIDRef = '6'; // 반전 ID
                        if (color === '#000000') color = '#FFFFFF';
                    } else {
                        shadeColor = `#${fill}`;
                        // 배경색용 borderFill 생성
                        borderFillIDRef = this.getOrCreateHighlightBorderFill(shadeColor);
                    }
                }
            }

            // Border (Box Text)
            const bdr = rPr.getElementsByTagNameNS(this.ns.w, 'bdr')[0];
            if (bdr) {
                borderFillIDRef = '5'; // Box ID
            }

            // Font (rFonts) - 언어별 폰트 분리
            const rFonts = rPr.getElementsByTagNameNS(this.ns.w, 'rFonts')[0];
            if (rFonts) {
                const eastAsia = rFonts.getAttribute('w:eastAsia');
                const ascii = rFonts.getAttribute('w:ascii');
                if (eastAsia) fontFace = eastAsia;
                else if (ascii) fontFace = ascii;
                // rFonts 엘리먼트를 저장해서 나중에 언어별 등록에 사용
                directRFonts = rFonts;
            }
        }

        // 2. Style Inheritance (if not direct)
        if (!hasDirectSize && paraStyleId) {
            let currentStyleId = paraStyleId;
            while (currentStyleId) {
                const style = this.docStyles.get(currentStyleId);
                if (!style) break;

                if (style.rPr && style.rPr.sz) {
                    const val = parseInt(style.rPr.sz);
                    if (!isNaN(val)) {
                        fontSize = val * 50;
                        break;
                    }
                }
                currentStyleId = style.basedOn;
            }
        }

        // 3. 스타일 체인에서 글자 색상 상속
        if (color === '#000000' && !isHyperlink && paraStyleId) {
            let currentStyleId = paraStyleId;
            while (currentStyleId) {
                const style = this.docStyles.get(currentStyleId);
                if (!style) break;

                if (style.rPr && style.rPr.color && style.rPr.color !== 'auto') {
                    color = `#${style.rPr.color}`;
                    break;
                }
                currentStyleId = style.basedOn;
            }
        }

        // 언어별 폰트 등록
        const fontIds = this.registerFontsFromRFonts(directRFonts);
        const height = fontSize;

        const key = `${bold}_${italic}_${underline}_${strike}_${supscript}_${subscript}_${color}_${shadeColor}_${borderFillIDRef}_${height}_${fontIds.hangulId}_${fontIds.latinId}`;

        for (const [id, cs] of this.charShapes) {
            if (cs._key === key) return id;
        }

        const newId = String(this.charShapes.size);
        this.charShapes.set(newId, {
            id: newId,
            height: String(height),
            textColor: color,
            shadeColor: shadeColor,
            borderFillIDRef: borderFillIDRef,
            bold, italic, underline, strike,
            supscript, subscript,
            // 언어별 폰트 ID 저장
            fontIds: fontIds,
            fontId: fontIds.hangulId, // 하위 호환
            _key: key
        });

        return newId;
    }
    getParagraphFontSize(para) {
        // 1. 직접 Run에서 폰트 크기 확인
        const runs = para.getElementsByTagNameNS(this.ns.w, 'r');
        for (const run of runs) {
            const rPr = run.getElementsByTagNameNS(this.ns.w, 'rPr')[0];
            if (rPr) {
                const sz = rPr.getElementsByTagNameNS(this.ns.w, 'sz')[0];
                if (sz) {
                    const val = sz.getAttribute('w:val');
                    if (val) {
                        return parseInt(val) / 2;
                    }
                }
            }
        }

        // 2. 문단 스타일에서 폰트 크기 상속
        const pPr = para.getElementsByTagNameNS(this.ns.w, 'pPr')[0];
        if (pPr) {
            // 문단 rPr에서 확인
            const pRPr = pPr.getElementsByTagNameNS(this.ns.w, 'rPr')[0];
            if (pRPr) {
                const sz = pRPr.getElementsByTagNameNS(this.ns.w, 'sz')[0];
                if (sz) {
                    const val = sz.getAttribute('w:val');
                    if (val) return parseInt(val) / 2;
                }
            }

            // 스타일 체인에서 상속
            const pStyle = pPr.getElementsByTagNameNS(this.ns.w, 'pStyle')[0];
            if (pStyle) {
                let currentStyleId = pStyle.getAttribute('w:val');
                while (currentStyleId) {
                    const style = this.docStyles.get(currentStyleId);
                    if (!style) break;

                    if (style.rPr && style.rPr.sz) {
                        const val = parseInt(style.rPr.sz);
                        if (!isNaN(val)) return val / 2;
                    }
                    currentStyleId = style.basedOn;
                }
            }
        }

        return 10; // 기본값 10pt
    }

    convertDrawing(drawing) {
        // 1. Pictures (blip)
        // 수정된 부분: 네임스페이스의 방해를 받지 않도록 모든 요소 중 localName이 'blip'인 것을 필터링
        const allElements = drawing.getElementsByTagName('*');
        const blips = Array.from(allElements).filter(el => el.localName === 'blip');

        // 결과 확인용 로그 (확인 후 삭제하셔도 됩니다)
        console.log(`✅ 찾은 이미지(blip) 개수: ${blips.length}`);
        if (blips.length > 0) {
            const embedId = blips[0].getAttributeNS(this.ns.r, 'embed') || blips[0].getAttribute('r:embed') || blips[0].getAttribute('embed');
            console.log(`- 추출된 이미지 연결 ID: ${embedId}`);

            return this.convertPicture(drawing, blips[0]);
        }

        // 2. Shapes (wps:wsp)
        // DOCX shapes are complex. We will approximate them as rectangles with checking allow overlap etc.
        const textboxes = drawing.getElementsByTagNameNS('http://schemas.microsoft.com/office/word/2010/wordprocessingShape', 'txbx');
        let textContent = '';
        if (textboxes.length > 0) {
            const paras = textboxes[0].getElementsByTagNameNS(this.ns.w, 'p');
            for (const p of paras) {
                const tNodes = p.getElementsByTagNameNS(this.ns.w, 't');
                for (const t of tNodes) textContent += t.textContent + '\n';
            }
        }

        if (textContent) {
            return `<hp:run charPrIDRef="0"><hp:t>[텍스트 상자: ${this.escapeXml(textContent.trim())}]</hp:t></hp:run>`;
        }

        // 3. Charts (c:chart)
        const charts = drawing.getElementsByTagName('c:chart');
        if (charts.length > 0 || drawing.textContent?.includes('chart')) {
            return `<hp:run charPrIDRef="0"><hp:t>[차트: 변환되지 않음]</hp:t></hp:run>`;
        }

        return '';
    }

    convertPicture(drawing, blip) {
        // JS 원본과 동일: getAttributeNS로 r:embed 접근, 폴백 추가
        const embed = blip.getAttributeNS(this.ns.r, 'embed') || blip.getAttribute('r:embed') || blip.getAttribute('embed');
        if (!embed) return '';
        const imgInfo = this.images.get(embed);
        console.log(this.images);
        if (!imgInfo) return '';

        const picId = this.tableIdCounter++;

        // 이미지 크기 추출 (EMU → HWPUNIT)
        let width = 27577; // 기본값
        let height = 19913;

        const extents = drawing.getElementsByTagNameNS(this.ns.wp, 'extent');
        if (extents.length > 0) {
            const cxAttr = extents[0].getAttribute('cx');
            const cyAttr = extents[0].getAttribute('cy');
            if (cxAttr && cyAttr) {
                width = Math.round(parseInt(cxAttr) * this.HWPUNIT_PER_INCH / this.EMU_PER_INCH);
                height = Math.round(parseInt(cyAttr) * this.HWPUNIT_PER_INCH / this.EMU_PER_INCH);
            }
        }

        const instId = Math.floor(Math.random() * 2000000000);

        // 인라인 vs 플로팅 판별
        const isInline = drawing.getElementsByTagNameNS(this.ns.wp, 'inline').length > 0;
        const anchor = drawing.getElementsByTagNameNS(this.ns.wp, 'anchor')[0];

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
            const wrapSquare = anchor.getElementsByTagNameNS(this.ns.wp, 'wrapSquare')[0];
            const wrapTight = anchor.getElementsByTagNameNS(this.ns.wp, 'wrapTight')[0];
            const wrapThrough = anchor.getElementsByTagNameNS(this.ns.wp, 'wrapThrough')[0];
            const wrapNone = anchor.getElementsByTagNameNS(this.ns.wp, 'wrapNone')[0];
            const wrapTopBottom = anchor.getElementsByTagNameNS(this.ns.wp, 'wrapTopAndBottom')[0];

            if (wrapSquare) textWrap = 'SQUARE';
            else if (wrapTight || wrapThrough) textWrap = 'TIGHT';
            else if (wrapNone) textWrap = 'BEHIND_TEXT';
            else if (wrapTopBottom) textWrap = 'TOP_AND_BOTTOM';

            // 수직 위치 파싱
            const posV = anchor.getElementsByTagNameNS(this.ns.wp, 'positionV')[0];
            if (posV) {
                const relFrom = posV.getAttribute('relativeFrom');
                if (relFrom === 'page') vertRelTo = 'PAGE';
                else if (relFrom === 'paragraph') vertRelTo = 'PARA';
                else if (relFrom === 'margin') vertRelTo = 'PAPER';

                const posOffset = posV.getElementsByTagNameNS(this.ns.wp, 'posOffset')[0];
                if (posOffset) {
                    // EMU → HWPUNIT
                    vertOffset = String(Math.round(parseInt(posOffset.textContent || '0') * this.HWPUNIT_PER_INCH / this.EMU_PER_INCH));
                }
                const align = posV.getElementsByTagNameNS(this.ns.wp, 'align')[0];
                if (align) {
                    const val = align.textContent;
                    if (val === 'center') vertAlign = 'CENTER';
                    else if (val === 'bottom') vertAlign = 'BOTTOM';
                    else if (val === 'top') vertAlign = 'TOP';
                }
            }

            // 수평 위치 파싱
            const posH = anchor.getElementsByTagNameNS(this.ns.wp, 'positionH')[0];
            if (posH) {
                const relFrom = posH.getAttribute('relativeFrom');
                if (relFrom === 'page') horzRelTo = 'PAGE';
                else if (relFrom === 'column') horzRelTo = 'COLUMN';
                else if (relFrom === 'margin') horzRelTo = 'PAPER';

                const posOffset = posH.getElementsByTagNameNS(this.ns.wp, 'posOffset')[0];
                if (posOffset) {
                    horzOffset = String(Math.round(parseInt(posOffset.textContent || '0') * this.HWPUNIT_PER_INCH / this.EMU_PER_INCH));
                }
                const align = posH.getElementsByTagNameNS(this.ns.wp, 'align')[0];
                if (align) {
                    const val = align.textContent;
                    if (val === 'center') horzAlign = 'CENTER';
                    else if (val === 'right') horzAlign = 'RIGHT';
                    else if (val === 'left') horzAlign = 'LEFT';
                }
            }

            // 외부 여백 (effectExtent)
            const effectExtent = anchor.getElementsByTagNameNS(this.ns.wp, 'effectExtent')[0];
            if (effectExtent) {
                const emuToHwp = (val) => String(Math.round(parseInt(val || '0') * this.HWPUNIT_PER_INCH / this.EMU_PER_INCH));
                outMarginLeft = emuToHwp(effectExtent.getAttribute('l'));
                outMarginRight = emuToHwp(effectExtent.getAttribute('r'));
                outMarginTop = emuToHwp(effectExtent.getAttribute('t'));
                outMarginBottom = emuToHwp(effectExtent.getAttribute('b'));
            }
        }

        // HWPX 이미지 XML 생성
        let xml = `<hp:run charPrIDRef="0">`;
        xml += `<hp:pic id="${picId}" zOrder="0" numberingType="NONE" textWrap="${textWrap}" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="${instId}" reverse="0">`;
        xml += `<hp:offset x="0" y="0"/>`;
        xml += `<hp:orgSz width="${width}" height="${height}"/>`;
        xml += `<hp:curSz width="${width}" height="${height}"/>`;
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


    convertTable(table) {
        const tableId = this.tableIdCounter++;

        // 테이블 속성 (tblPr)
        // 테이블 속성 (tblPr)
        const tblPr = table.getElementsByTagNameNS(this.ns.w, 'tblPr')[0];
        let tblBorders = null;
        let defaultMargins = { left: 283, right: 283, top: 283, bottom: 283 }; // Default ~1mm

        if (tblPr) {
            tblBorders = tblPr.getElementsByTagNameNS(this.ns.w, 'tblBorders')[0];

            // Table-level cell margins (tblCellMar)
            const tblCellMar = tblPr.getElementsByTagNameNS(this.ns.w, 'tblCellMar')[0];
            if (tblCellMar) {
                const parseMar = (tag, def) => {
                    const m = tblCellMar.getElementsByTagNameNS(this.ns.w, tag)[0];
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

        // Grid (컬럼 너비)
        const tblGrid = table.getElementsByTagNameNS(this.ns.w, 'tblGrid')[0];
        const gridCols = tblGrid ? tblGrid.getElementsByTagNameNS(this.ns.w, 'gridCol') : [];
        const colWidths = [];
        let totalWidth = 0;
        for (const col of gridCols) {
            const w = parseInt(col.getAttribute('w:w') || '0');
            // Twips to HWPUNIT
            const wHwp = w * 5;
            colWidths.push(wHwp);
            totalWidth += wHwp;
        }

        // Rows
        const rows = table.getElementsByTagNameNS(this.ns.w, 'tr');
        const rowCount = rows.length;
        const colCount = colWidths.length || 1;

        // Row/Col Span Pre-calculation (vMerge, gridSpan)
        const cellMap = []; // [row][col] -> { span: boolean, rowSpan: 1, colSpan: 1, ... }
        for (let r = 0; r < rowCount; r++) {
            cellMap[r] = [];
            for (let c = 0; c < colCount; c++) {
                cellMap[r][c] = { occupied: false, rowSpan: 1, colSpan: 1, startRow: r, startCol: c };
            }
        }

        // 1. GridSpan & vMerge parsing
        for (let r = 0; r < rowCount; r++) {
            const row = rows[r];
            const cells = row.getElementsByTagNameNS(this.ns.w, 'tc');
            let colIndex = 0;

            for (const cell of cells) {
                // Find next unoccupied column
                while (colIndex < colCount && cellMap[r][colIndex].occupied) {
                    colIndex++;
                }
                if (colIndex >= colCount) break;

                const tcPr = cell.getElementsByTagNameNS(this.ns.w, 'tcPr')[0];
                let gridSpan = 1;
                let vMerge = null;

                if (tcPr) {
                    const gridSpanElem = tcPr.getElementsByTagNameNS(this.ns.w, 'gridSpan')[0];
                    if (gridSpanElem) gridSpan = parseInt(gridSpanElem.getAttribute('w:val') || '1');

                    const vMergeElem = tcPr.getElementsByTagNameNS(this.ns.w, 'vMerge')[0];
                    if (vMergeElem) {
                        vMerge = vMergeElem.getAttribute('w:val') || 'continue';
                    }
                }

                // Mark gridSpan
                for (let i = 0; i < gridSpan; i++) {
                    if (colIndex + i < colCount) {
                        cellMap[r][colIndex + i].occupied = true;
                        cellMap[r][colIndex + i].cell = cell; // Reference to logic cell
                        // Main cell info
                        if (i === 0) {
                            cellMap[r][colIndex].colSpan = gridSpan;
                            cellMap[r][colIndex].vMerge = vMerge;
                        }
                    }
                }

                colIndex += gridSpan;
            }
        }

        // 2. Resolve vMerge (RowSpans)
        for (let c = 0; c < colCount; c++) {
            for (let r = 0; r < rowCount; r++) {
                const info = cellMap[r][c];
                // If it's a restart or standard cell, check for continuations below
                if (info.cell && (info.vMerge === 'restart' || !info.vMerge)) {
                    let rowSpan = 1;
                    for (let nextR = r + 1; nextR < rowCount; nextR++) {
                        const nextInfo = cellMap[nextR][c];
                        if (nextInfo.cell && nextInfo.vMerge === 'continue') {
                            rowSpan++;
                            // Mark as merged part (don't output)
                            nextInfo.isMergedBelow = true;
                        } else {
                            break;
                        }
                    }
                    info.rowSpan = rowSpan;
                }
            }
        }

        // Safety: Ensure Total Width is valid (User Requirement: Width > 0)
        if (totalWidth === 0) {
            totalWidth = this.TEXT_WIDTH;
            // Distribute width evenly if colWidths are missing
            if (colWidths.length === 0 && colCount > 0) {
                const avgWidth = Math.floor(totalWidth / colCount);
                for (let i = 0; i < colCount; i++) colWidths.push(avgWidth);
            }
        }

        // Pre-calculate Table Height
        let tableHeight = 0;
        for (let r = 0; r < rowCount; r++) {
            const row = rows[r];
            let h = 0;
            const trPr = row.getElementsByTagNameNS(this.ns.w, 'trPr')[0];
            if (trPr) {
                const trHeightElem = trPr.getElementsByTagNameNS(this.ns.w, 'trHeight')[0];
                if (trHeightElem) {
                    const val = parseInt(trHeightElem.getAttribute('w:val') || '0');
                    h = val * 5; // Twips to HWPUNIT
                }
            }
            tableHeight += (h > 0 ? h : 1500); // 0이면 기본 높이(약 2000) 가산
        }

        // [HWP 5.0 Spec] 표 컨트롤(0x0B)에 해당하는 XML 구조
        // hp:tbl은 문단 텍스트 내의 컨트롤 문자(0x0B) 역할을 함
        // User Rule: treatAsChar="0", Empty Caption Removal
        // Wrap in hp:p
        let xml = `<hp:p id="${this.idCounter++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:tbl id="${tableId}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="${rowCount}" colCnt="${colCount}" cellSpacing="0" borderFillIDRef="3" noAdjust="0">`;
        xml += `<hp:sz width="${totalWidth}" widthRelTo="ABSOLUTE" height="${tableHeight}" heightRelTo="ABSOLUTE" protect="0"/>`;

        xml += `<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>`;
        xml += `<hp:outMargin left="0" right="0" top="0" bottom="0"/>`;
        // <hp:caption> Deleted per Strict Rule (Don't render empty caption)
        xml += `<hp:inMargin left="540" right="540" top="0" bottom="0"/>`;


        // Output rows
        for (let r = 0; r < rowCount; r++) {
            xml += `<hp:tr>`;

            // Row Height
            const row = rows[r];
            let trHeight = 0; // 0 for auto
            const trPr = row.getElementsByTagNameNS(this.ns.w, 'trPr')[0];
            if (trPr) {
                const trHeightElem = trPr.getElementsByTagNameNS(this.ns.w, 'trHeight')[0];
                if (trHeightElem) {
                    const val = parseInt(trHeightElem.getAttribute('w:val') || '0');
                    // Twips to HWPUNIT
                    trHeight = val * 5;
                }
            }

            let c = 0;
            while (c < colCount) {
                const info = cellMap[r][c];

                if (info.isMergedBelow) {
                    c++;
                    continue;
                }

                const cell = info.cell;
                if (!cell) {
                    // Empty/Missing cell fill
                    // Width calculation
                    let cellWidth = 0;
                    for (let k = 0; k < info.colSpan; k++) {
                        cellWidth += colWidths[c + k] || 0;
                    }
                    xml += `<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="4">`;
                    xml += `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="TOP" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">`;
                    xml += `<hp:p id="${this.idCounter++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t></hp:t></hp:run><hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="0" horzpos="0" horzsize="${cellWidth}" flags="393216"/></hp:linesegarray></hp:p>`;
                    xml += `</hp:subList>`;
                    xml += `<hp:cellAddr colAddr="${c}" rowAddr="${r}"/>`;
                    xml += `<hp:cellSpan colSpan="${info.colSpan}" rowSpan="${info.rowSpan}"/>`;
                    xml += `<hp:cellSz width="${cellWidth}" height="0"/>`;
                    xml += `<hp:cellMargin left="255" right="255" top="141" bottom="141"/>`;
                    xml += `</hp:tc>`;
                    c++;
                    continue;
                }

                // Width calculation
                let cellWidth = 0;
                for (let k = 0; k < info.colSpan; k++) {
                    cellWidth += colWidths[c + k] || 0;
                }

                const tcPr = cell.getElementsByTagNameNS(this.ns.w, 'tcPr')[0];
                const borderFill = this.getCellBorderFill(tcPr, tblBorders, r, c, rowCount, colCount, info.rowSpan, info.colSpan);

                // Cell Margins
                let margins = { ...defaultMargins }; // Start with table defaults

                // Parse tcMar if exists (Override)
                if (tcPr) {
                    const tcMar = tcPr.getElementsByTagNameNS(this.ns.w, 'tcMar')[0];
                    if (tcMar) {
                        const parseMar = (tag) => {
                            const mTag = tcMar.getElementsByTagNameNS(this.ns.w, tag)[0];
                            if (mTag) {
                                const w = parseInt(mTag.getAttribute('w:w') || '0');
                                return w * 5;
                            }
                            return null;
                        };

                        const l = parseMar('left'); if (l !== null) margins.left = l;
                        const rVal = parseMar('right'); if (rVal !== null) margins.right = rVal;
                        const t = parseMar('top'); if (t !== null) margins.top = t;
                        const b = parseMar('bottom'); if (b !== null) margins.bottom = b;
                    }
                }

                // hasMargin="0" (참조 기본값), 커스텀 마진이 있으면 "1"
                const hasTcMar = (tcPr && tcPr.getElementsByTagNameNS(this.ns.w, 'tcMar')[0]) ? '1' : '0';
                xml += `<hp:tc name="" header="0" hasMargin="${hasTcMar}" protect="0" editable="0" dirty="0" borderFillIDRef="${borderFill.id}">`;

                // Vertical Alignment (tcPr > vAlign)
                let vertAlign = 'TOP'; // Default
                if (tcPr) {
                    const vAlign = tcPr.getElementsByTagNameNS(this.ns.w, 'vAlign')[0];
                    if (vAlign) {
                        const val = vAlign.getAttribute('w:val');
                        if (val === 'center') vertAlign = 'CENTER';
                        else if (val === 'bottom') vertAlign = 'BOTTOM';
                    }
                }

                // [HWPX 참조] 순서: subList → cellAddr → cellSpan → cellSz → cellMargin
                xml += `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="${vertAlign}" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">`;

                // Cell Content (문단 + 중첩 테이블 모두 처리)
                let hasContent = false;
                for (const childNode of cell.childNodes) {
                    if (childNode.nodeName === 'w:p') {
                        xml += this.convertParagraph(childNode);
                        hasContent = true;
                    } else if (childNode.nodeName === 'w:tbl') {
                        // 중첩 테이블 — 테이블 안의 테이블
                        xml += this.convertTable(childNode);
                        hasContent = true;
                    }
                }
                if (!hasContent) {
                    // 빈 셀 안전장치
                    xml += `<hp:p id="${this.idCounter++}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:t></hp:t></hp:run><hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="0" horzpos="0" horzsize="${cellWidth}" flags="393216"/></hp:linesegarray></hp:p>`;
                }

                xml += `</hp:subList>`;

                // 셀 속성 (subList 뒤에 배치 — 참조 구조)
                xml += `<hp:cellAddr colAddr="${c}" rowAddr="${r}"/>`;
                xml += `<hp:cellSpan colSpan="${info.colSpan}" rowSpan="${info.rowSpan}"/>`;
                xml += `<hp:cellSz width="${cellWidth}" height="0"/>`; // height=0: 자동 높이
                xml += `<hp:cellMargin left="${margins.left}" right="${margins.right}" top="${margins.top}" bottom="${margins.bottom}"/>`;

                xml += `</hp:tc>`;

                c += info.colSpan;
            }
            xml += `</hp:tr>`;
        }

        xml += `</hp:tbl></hp:run>`;

        // Add lineseg for the table paragraph
        xml += `<hp:linesegarray><hp:lineseg textpos="0" vertpos="${this.currentVertPos}" vertsize="${tableHeight}" textheight="${tableHeight}" baseline="${tableHeight}" spacing="0" horzpos="0" horzsize="${totalWidth}" flags="393216"/></hp:linesegarray>`;
        xml += `</hp:p>`;

        this.currentVertPos += tableHeight;

        return xml;
    }

    createContentHpf() {
        const now = new Date();
        const isoDate = now.toISOString();
        // 참조 형식에 맞춘 content.hpf 생성 - 모든 네임스페이스 포함
        let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>`;
        xml += `<opf:package xmlns:ha="${this.hwpxNs.ha}" xmlns:hp="${this.hwpxNs.hp}" xmlns:hp10="${this.hwpxNs.hp10}" xmlns:hs="${this.hwpxNs.hs}" xmlns:hc="${this.hwpxNs.hc}" xmlns:hh="${this.hwpxNs.hh}" xmlns:hhs="${this.hwpxNs.hhs}" xmlns:hm="${this.hwpxNs.hm}" xmlns:hpf="${this.hwpxNs.hpf}" xmlns:dc="${this.hwpxNs.dc}" xmlns:opf="${this.hwpxNs.opf}" xmlns:ooxmlchart="${this.hwpxNs.ooxmlchart}" xmlns:epub="${this.hwpxNs.epub}" xmlns:config="${this.hwpxNs.config}" version="" unique-identifier="" id="">`;

        // 메타데이터 - opf: 네임스페이스 사용 (참조 형식)
        xml += `<opf:metadata>`;
        xml += `<opf:title>${this.escapeXml(this.metadata.title || '문서')}</opf:title>`;
        xml += `<opf:language>ko</opf:language>`;
        if (this.metadata.creator) {
            xml += `<opf:meta name="creator" content="text">${this.escapeXml(this.metadata.creator)}</opf:meta>`;
        }
        xml += `<opf:meta name="subject" content="text">${this.escapeXml(this.metadata.subject || '')}</opf:meta>`;
        if (this.metadata.description) {
            xml += `<opf:meta name="description" content="text">${this.escapeXml(this.metadata.description)}</opf:meta>`;
        }
        xml += `<opf:meta name="CreatedDate" content="text">${isoDate}</opf:meta>`;
        xml += `<opf:meta name="ModifiedDate" content="text">${isoDate}</opf:meta>`;
        xml += `</opf:metadata>`;

        // 매니페스트
        xml += `<opf:manifest>`;

        // 이미지 매니페스트 (이미지가 먼저)
        for (const [rId, imgInfo] of this.images) {
            const mediaType = imgInfo.ext === 'png' ? 'image/png' : imgInfo.ext === 'jpg' || imgInfo.ext === 'jpeg' ? 'image/jpeg' : imgInfo.ext === 'gif' ? 'image/gif' : 'image/png';
            xml += `<opf:item id="${imgInfo.manifestId}" href="${imgInfo.path}" media-type="${mediaType}" isEmbeded="1"/>`;
        }

        xml += `<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>`;
        xml += `<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>`;
        xml += `<opf:item id="settings" href="settings.xml" media-type="application/xml"/>`;
        xml += `</opf:manifest>`;

        xml += `<opf:spine>`;
        xml += `<opf:itemref idref="header"/>`;
        xml += `<opf:itemref idref="section0"/>`;
        xml += `</opf:spine>`;
        xml += `</opf:package>`;
        return xml;
    }

    createContainer() {
        // 참조 형식에 맞춘 단일 줄 compact XML
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/><ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/><ocf:rootfile full-path="META-INF/container.rdf" media-type="application/rdf+xml"/></ocf:rootfiles></ocf:container>`;
    }

    createContainerhdf() {
        // container.rdf - 참조 형식에 맞춘 compact XML
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/header.xml"/></rdf:Description><rdf:Description rdf:about="Contents/header.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/></rdf:Description><rdf:Description rdf:about=""><ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" rdf:resource="Contents/section0.xml"/></rdf:Description><rdf:Description rdf:about="Contents/section0.xml"><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/></rdf:Description><rdf:Description rdf:about=""><rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"/></rdf:Description></rdf:RDF>`;
    }

    createManifest() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><manifest xmlns="${this.hwpxNs.odf}"/>`;
    }

    async packageHwpx(hwpxData) {
        const zip = new JSZip();

        // mimetype (압축 안 함 - PDF 표 2 참고)
        zip.file('mimetype', hwpxData.mimetype, { compression: 'STORE' });

        // XML 파일들
        zip.file('version.xml', hwpxData.version);
        zip.file('settings.xml', hwpxData.settings);
        zip.file('Contents/header.xml', hwpxData.header);
        zip.file('Contents/section0.xml', hwpxData.section0);
        zip.file('Contents/content.hpf', hwpxData.contentHpf);
        zip.file('META-INF/container.xml', hwpxData.container);
        zip.file('META-INF/container.rdf', hwpxData.containerRdf);
        zip.file('META-INF/manifest.xml', hwpxData.manifest);

        // Preview 폴더
        zip.file('Preview/PrvText.txt', '문서 미리보기');

        // 이미지들
        for (const [rId, imgInfo] of this.images) {
            zip.file(imgInfo.path, imgInfo.data);
        }

        console.log(`✅ Packaged HWPX with ${this.images.size} images`);

        return await zip.generateAsync({ type: 'blob' });
    }

    // 특정 언어에 폰트 등록
    registerFontForLang(lang, face) {
        const map = this.langFontFaces[lang];
        if (!map) return 0;
        if (!map.has(face)) {
            const id = map.size;
            map.set(face, id);
        }
        return map.get(face);
    }

    // 모든 언어에 폰트 등록 (하위 호환)
    registerFont(face) {
        const langs = ['HANGUL', 'LATIN', 'HANJA', 'JAPANESE', 'OTHER', 'SYMBOL', 'USER'];
        let id = 0;
        for (const lang of langs) {
            id = this.registerFontForLang(lang, face);
        }
        return id;
    }

    // DOCX rFonts에서 언어별 폰트 등록 및 ID 반환
    registerFontsFromRFonts(rFonts) {
        let hangulFont = '함초롬바탕';
        let latinFont = '함초롬바탕';

        if (rFonts) {
            const eastAsia = rFonts.getAttribute('w:eastAsia');
            const ascii = rFonts.getAttribute('w:ascii');
            const hAnsi = rFonts.getAttribute('w:hAnsi');
            const cs = rFonts.getAttribute('w:cs');

            if (eastAsia) hangulFont = eastAsia;
            if (ascii) latinFont = ascii;
            else if (hAnsi) latinFont = hAnsi;
        }

        const hangulId = this.registerFontForLang('HANGUL', hangulFont);
        const latinId = this.registerFontForLang('LATIN', latinFont);
        const hanjaId = this.registerFontForLang('HANJA', hangulFont);
        const japaneseId = this.registerFontForLang('JAPANESE', hangulFont);
        const otherId = this.registerFontForLang('OTHER', latinFont);
        const symbolId = this.registerFontForLang('SYMBOL', latinFont);
        const userId = this.registerFontForLang('USER', hangulFont);

        return { hangulId, latinId, hanjaId, japaneseId, otherId, symbolId, userId };
    }

    // 고정폭 글꼴 감지
    isMonospaceFont(face) {
        const monoFonts = [
            'Courier New', 'Consolas', 'Lucida Console', 'Monaco',
            'Menlo', 'monospace', 'Ubuntu Mono', 'Source Code Pro',
            'Fira Code', 'JetBrains Mono', 'D2Coding', 'NanumGothicCoding'
        ];
        return monoFonts.some(m => face.toLowerCase().includes(m.toLowerCase()));
    }

    // 폰트 패밀리 타입 판별
    getFontFamilyType(face) {
        if (this.isMonospaceFont(face)) return 'FCAT_FIXED';
        // 고딕 계열
        const gothicFonts = ['Gothic', '고딕', '돋움', 'Arial', 'Helvetica', 'Verdana', 'Tahoma', 'Sans'];
        if (gothicFonts.some(g => face.includes(g))) return 'FCAT_GOTHIC';
        // 명조/바탕 계열
        const serifFonts = ['바탕', '명조', 'Batang', 'Times', 'Serif', 'Georgia'];
        if (serifFonts.some(s => face.includes(s))) return 'FCAT_MYEONGJO';
        return 'FCAT_UNKNOWN';
    }

    createFontFaces() {
        const langs = ['HANGUL', 'LATIN', 'HANJA', 'JAPANESE', 'OTHER', 'SYMBOL', 'USER'];
        // itemCnt는 언어 카테고리 수 (7)
        let xml = `<hh:fontfaces itemCnt="${langs.length}">`;

        for (const lang of langs) {
            const fontMap = this.langFontFaces[lang];
            xml += `<hh:fontface lang="${lang}" fontCnt="${fontMap.size}">`;
            for (const [face, id] of fontMap) {
                const familyType = this.getFontFamilyType(face);
                xml += `<hh:font id="${id}" face="${face}" type="TTF" isEmbedded="0">`;
                xml += `<hh:typeInfo familyType="${familyType}" weight="0" proportion="0" contrast="0" strokeVariation="0" armStyle="0" letterform="0" midline="0" xHeight="0"/>`;
                xml += `</hh:font>`;
            }
            xml += `</hh:fontface>`;
        }
        xml += `</hh:fontfaces>`;
        return xml;
    }


    extractText(para) {
        let text = '';
        const texts = para.getElementsByTagNameNS(this.ns.w, 't');
        for (const t of texts) {
            text += t.textContent;
        }
        return text;
    }

    isParagraphBold(para) {
        // Check pPr default run properties
        const pPr = para.getElementsByTagNameNS(this.ns.w, 'pPr')[0];
        if (pPr) {
            const rPr = pPr.getElementsByTagNameNS(this.ns.w, 'rPr')[0];
            if (rPr && rPr.getElementsByTagNameNS(this.ns.w, 'b').length > 0) return true;
        }
        // Check first run as heuristic
        const runs = para.getElementsByTagNameNS(this.ns.w, 'r');
        if (runs.length > 0) {
            const rPr = runs[0].getElementsByTagNameNS(this.ns.w, 'rPr')[0];
            if (rPr && rPr.getElementsByTagNameNS(this.ns.w, 'b').length > 0) return true;
        }
        return false;
    }

    convertParagraph(para) {
        const pPr = para.getElementsByTagNameNS(this.ns.w, 'pPr')[0];
        let paraStyleId = '0'; // Default style ID

        if (pPr) {
            const pStyle = pPr.getElementsByTagNameNS(this.ns.w, 'pStyle')[0];
            if (pStyle) paraStyleId = pStyle.getAttribute('w:val');
        }

        const id = this.getParaShapeId(para);
        const paraShape = this.paraShapes.get(id);

        let xml = `<hp:p id="${this.idCounter++}" paraPrIDRef="${id}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">`;

        // Run Processing
        const runs = para.getElementsByTagNameNS(this.ns.w, 'r');
        if (runs.length === 0) {
            // Empty Paragraph
            // We should still pass paraStyleId to getCharShapeId for the empty run properties if needed, 
            // but empty run usually just needs a default char shape.
            // Let's pass it anyway to be safe for default font size in empty para.
            const charPrIDRef = this.getCharShapeId(null, undefined, false, paraStyleId);
            xml += `<hp:run charPrIDRef="${charPrIDRef}"><hp:t></hp:t></hp:run>`;
        } else {
            for (const run of runs) {
                const rPr = run.getElementsByTagNameNS(this.ns.w, 'rPr')[0];
                const texts = run.getElementsByTagNameNS(this.ns.w, 't');
                const tabs = run.getElementsByTagNameNS(this.ns.w, 'tab');
                const brs = run.getElementsByTagNameNS(this.ns.w, 'br');
                const images = run.getElementsByTagNameNS(this.ns.w, 'drawing');
                const objects = run.getElementsByTagNameNS(this.ns.w, 'object');
                const picts = run.getElementsByTagNameNS(this.ns.w, 'pict'); // Legacy VML

                // Hyperlink check (simplified)
                const isHyperlink = false;

                const charPrIDRef = this.getCharShapeId(rPr, undefined, isHyperlink, paraStyleId);

                // ... content processing ...
                // This part was missing in the provided snippet, so I'm keeping the original logic here.
                // If the intention was to replace the entire convertParagraph, the snippet was incomplete.
                // Assuming the user wants to integrate the charPrIDRef logic into the existing run processing.

                let hasTextContent = false;
                xml += `<hp:run charPrIDRef="${charPrIDRef}">`;

                for (const t of texts) {
                    const textContent = t.textContent;
                    if (textContent.length > 0) {
                        xml += `<hp:t>${this.escapeXml(textContent)}</hp:t>`;
                        hasTextContent = true;
                    }
                }

                for (const tab of tabs) {
                    xml += `<hp:tab/>`;
                }

                for (const br of brs) {
                    const type = br.getAttribute('w:type');
                    if (type === 'page') {
                        xml += `<hp:lineSegOnly/>`; // Placeholder for page break
                    } else {
                        xml += `<hp:lineBreak/>`;
                    }
                }

                xml += `</hp:run>`;

                // 이미지는 별도의 hp:run으로 처리 (중첩 방지)
                for (const img of images) {
                    xml += this.convertDrawing(img);
                }

                for (const obj of objects) {
                    // Handle embedded objects if necessary
                }

                for (const pict of picts) {
                    // Handle VML images if necessary
                }
            }
        }

        xml += `</hp:p>`;
        return xml;
    }

    getCellBorderFill(tcPr, tblBorders, r, c, rowCount, colCount, rowSpan = 1, colSpan = 1) {
        // Default borders
        const borders = {
            left: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            right: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            top: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            bottom: { type: 'SOLID', width: '0.1 mm', color: '#000000' },
            slash: { type: 'NONE', width: '0.1 mm', color: '#000000' },
            backSlash: { type: 'NONE', width: '0.1 mm', color: '#000000' }
        };

        const extract = (elem) => {
            if (!elem) return null;
            const val = elem.getAttribute('w:val');
            if (val === 'nil' || val === 'none') return { type: 'NONE', width: '0.1 mm', color: '#000000' };

            let type = 'SOLID';
            if (val === 'double') type = 'DOUBLE';
            else if (val === 'dashed') type = 'DASH';
            else if (val === 'dotted') type = 'DOT';
            else if (val === 'thick') type = 'SOLID'; // Bold line
            else if (val === 'thinThickSmallGap') type = 'DOUBLE'; // Approx
            // Add more types as needed

            // DOCX sz: 1/8 point
            // HWP Standard Thicknesses (mm): 0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0
            const sz = parseInt(elem.getAttribute('w:sz') || '4');
            const mm = (sz / 8) * 0.3528;

            // HWP Thickness snapping
            const presets = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0];
            let widthVal = presets[0];
            let minDiff = Math.abs(mm - presets[0]);

            for (let i = 1; i < presets.length; i++) {
                const diff = Math.abs(mm - presets[i]);
                if (diff < minDiff) {
                    minDiff = diff;
                    widthVal = presets[i];
                }
            }

            // Special case: if DOCX is very thick, max out at 5.0
            if (mm > 5.0) widthVal = 5.0;

            let color = '#000000';
            const colorAttr = elem.getAttribute('w:color');
            if (colorAttr && colorAttr !== 'auto') color = `#${colorAttr}`;

            return { type, width: `${widthVal} mm`, color };
        };

        const mapBorder = (source, side, targetSide) => {
            if (!source) return;
            const borderElem = source.getElementsByTagNameNS(this.ns.w, side)[0];
            if (borderElem) {
                const info = extract(borderElem);
                if (info) borders[targetSide] = info;
            }
        };

        // 1. Apply Table Level Borders (Smart Logic)
        if (tblBorders) {
            const insideH = tblBorders.getElementsByTagNameNS(this.ns.w, 'insideH')[0];
            const insideV = tblBorders.getElementsByTagNameNS(this.ns.w, 'insideV')[0];
            const defH = extract(insideH);
            const defV = extract(insideV);

            // Apply Inner Defaults
            // These apply to all internal borders, which can then be overridden by explicit outer borders or tcBorders
            if (defH) { borders.top = { ...defH }; borders.bottom = { ...defH }; }
            if (defV) { borders.left = { ...defV }; borders.right = { ...defV }; }

            // Apply Outer Borders based on position
            // Top Edge
            if (r === 0) {
                mapBorder(tblBorders, 'top', 'top');
            }
            // Bottom Edge
            if (r + rowSpan >= rowCount) {
                mapBorder(tblBorders, 'bottom', 'bottom');
            }
            // Left Edge
            if (c === 0) {
                mapBorder(tblBorders, 'left', 'left');
            }
            // Right Edge
            if (c + colSpan >= colCount) {
                mapBorder(tblBorders, 'right', 'right');
            }
        }

        // 2. Overwrite with explicit tcBorders
        if (tcPr) {
            const tcBorders = tcPr.getElementsByTagNameNS(this.ns.w, 'tcBorders')[0];
            if (tcBorders) {
                mapBorder(tcBorders, 'top', 'top');
                mapBorder(tcBorders, 'bottom', 'bottom');
                mapBorder(tcBorders, 'left', 'left');
                mapBorder(tcBorders, 'right', 'right');
                mapBorder(tcBorders, 'tl2br', 'slash');     // Diagonal /
                mapBorder(tcBorders, 'tr2bl', 'backSlash'); // Diagonal \
            }
        }

        // Background color (shading)
        let backColor = '#FFFFFF';
        if (tcPr) {
            const shd = tcPr.getElementsByTagNameNS(this.ns.w, 'shd')[0];
            if (shd) {
                const fill = shd.getAttribute('w:fill');
                if (fill && fill !== 'auto') backColor = `#${fill}`;
            }
        }

        // Generate Key & Cache
        // Fixed: Added backColor and diagonals to the key to prevent collisions
        const keyObj = {
            L: borders.left, R: borders.right, T: borders.top, B: borders.bottom,
            Sl: borders.slash, Bs: borders.backSlash,
            BG: backColor
        };
        const key = JSON.stringify(keyObj);

        for (const [id, bf] of this.borderFills) {
            if (id === '1' || id === '2' || id === '3' || id === '4') continue;
            if (bf._key === key) {
                return bf;
            }
        }

        const newId = String(this.borderFills.size + 1);
        const newBf = {
            id: newId,
            leftBorder: borders.left,
            rightBorder: borders.right,
            topBorder: borders.top,
            bottomBorder: borders.bottom,
            slash: borders.slash,
            backSlash: borders.backSlash,
            backColor: backColor,
            _key: key
        };

        this.borderFills.set(newId, newBf);
        return newBf;
    }

    escapeXml(text) {
        if (!text) return '';
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }
}

// Node.js 환경 지원
if (typeof module !== 'undefined' && module.exports) {
    module.exports = DocxToHwpxConverter;
}
