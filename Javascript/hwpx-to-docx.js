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

const JSZip = require('jszip');
const { DOMParser } = require('xmldom');
const fs = require('fs');
const path = require('path');

class HwpxToDocxConverter {
    constructor() {
        // HWPX 네임스페이스
        this.ns = {
            hp: 'http://www.hancom.co.kr/hwpml/2011/paragraph',
            hh: 'http://www.hancom.co.kr/hwpml/2011/head',
            hc: 'http://www.hancom.co.kr/hwpml/2011/core',
            hs: 'http://www.hancom.co.kr/hwpml/2011/section',
            ha: 'http://www.hancom.co.kr/hwpml/2011/app',
            opf: 'http://www.idpf.org/2007/opf/'
        };

        // DOCX 네임스페이스
        this.docxNs = {
            w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
            pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        };

        // 파싱된 데이터 저장소
        this.charProperties = new Map();  // id → charPr 정보
        this.paraProperties = new Map();  // id → paraPr 정보
        this.borderFills = new Map();     // id → borderFill 정보
        this.fontFaces = {};              // lang → [fonts]
        this.styles = new Map();          // id → style 정보
        this.numberings = new Map();      // id → numbering 정보
        this.images = new Map();          // binaryItemIDRef → {data, mediaType, ext}
        this.imageRels = [];              // 이미지 relationship 목록

        // 페이지 설정 기본값
        this.pageWidth = 11906;   // A4 width in twips
        this.pageHeight = 16838;  // A4 height in twips
        this.marginTop = 1440;
        this.marginBottom = 1440;
        this.marginLeft = 1800;
        this.marginRight = 1800;
        this.marginHeader = 851;
        this.marginFooter = 851;

        // 관계 ID 카운터
        this.relIdCounter = 1;
    }

    // ========================================
    // 메인 변환 진입점
    // ========================================
    async convert(hwpxInput) {
        let zip;
        if (typeof hwpxInput === 'string') {
            // 디렉토리 경로인 경우 (압축 해제된 HWPX)
            if (fs.existsSync(hwpxInput) && fs.statSync(hwpxInput).isDirectory()) {
                return await this.convertFromDirectory(hwpxInput);
            }
            // 파일 경로인 경우
            const data = fs.readFileSync(hwpxInput);
            zip = await JSZip.loadAsync(data);
        } else if (hwpxInput instanceof Buffer) {
            zip = await JSZip.loadAsync(hwpxInput);
        } else {
            zip = hwpxInput; // JSZip 인스턴스
        }
        return await this.convertFromZip(zip);
    }

    // 디렉토리에서 변환 (압축 해제된 HWPX)
    async convertFromDirectory(dirPath) {
        console.log(`HWPX 디렉토리에서 변환: ${dirPath}`);

        // content.hpf 파싱하여 이미지 매핑 가져오기
        const contentHpfPath = path.join(dirPath, 'Contents', 'content.hpf');
        if (fs.existsSync(contentHpfPath)) {
            const xml = fs.readFileSync(contentHpfPath, 'utf-8');
            this.parseContentHpf(xml);
        }

        // header.xml 파싱
        const headerPath = path.join(dirPath, 'Contents', 'header.xml');
        if (fs.existsSync(headerPath)) {
            const xml = fs.readFileSync(headerPath, 'utf-8');
            this.parseHeader(xml);
        }

        // 이미지 로드
        const binDataPath = path.join(dirPath, 'BinData');
        if (fs.existsSync(binDataPath)) {
            const files = fs.readdirSync(binDataPath);
            for (const file of files) {
                const filePath = path.join(binDataPath, file);
                const data = fs.readFileSync(filePath);
                const ext = path.extname(file).toLowerCase().replace('.', '');
                const id = path.basename(file, path.extname(file));
                const mediaType = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`;
                this.images.set(id, { data, mediaType, ext });
            }
        }

        // section0.xml 파싱 및 변환
        const sectionPath = path.join(dirPath, 'Contents', 'section0.xml');
        let bodyXml = '';
        if (fs.existsSync(sectionPath)) {
            const xml = fs.readFileSync(sectionPath, 'utf-8');
            bodyXml = this.convertSection(xml);
        }

        // DOCX 패키지 생성
        return await this.createDocxPackage(bodyXml);
    }

    // ZIP에서 변환
    async convertFromZip(zip) {
        console.log('HWPX ZIP에서 변환...');

        // content.hpf
        const contentHpf = zip.file('Contents/content.hpf');
        if (contentHpf) {
            this.parseContentHpf(await contentHpf.async('string'));
        }

        // header.xml
        const header = zip.file('Contents/header.xml');
        if (header) {
            this.parseHeader(await header.async('string'));
        }

        // 이미지 로드
        const binDataFolder = zip.folder('BinData');
        if (binDataFolder) {
            const imageFiles = [];
            binDataFolder.forEach((relativePath, file) => {
                if (!file.dir) imageFiles.push({ name: relativePath, file });
            });
            for (const { name, file } of imageFiles) {
                const data = await file.async('nodebuffer');
                const ext = path.extname(name).toLowerCase().replace('.', '');
                const id = path.basename(name, path.extname(name));
                const mediaType = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`;
                this.images.set(id, { data, mediaType, ext });
            }
        }

        // section0.xml
        const section = zip.file('Contents/section0.xml');
        let bodyXml = '';
        if (section) {
            bodyXml = this.convertSection(await section.async('string'));
        }

        return await this.createDocxPackage(bodyXml);
    }

    // ========================================
    // content.hpf 파싱
    // ========================================
    parseContentHpf(xml) {
        const doc = new DOMParser().parseFromString(xml, 'text/xml');
        // 이미지 아이템 매핑은 images Map에서 직접 처리
        console.log('content.hpf 파싱 완료');
    }

    // ========================================
    // header.xml 파싱
    // ========================================
    parseHeader(xml) {
        const doc = new DOMParser().parseFromString(xml, 'text/xml');

        // 1. fontfaces 파싱
        this.parseFontFaces(doc);

        // 2. borderFills 파싱
        this.parseBorderFills(doc);

        // 3. charProperties 파싱
        this.parseCharProperties(doc);

        // 4. paraProperties 파싱
        this.parseParaProperties(doc);

        // 5. styles 파싱
        this.parseStyles(doc);

        // 6. numberings 파싱
        this.parseNumberings(doc);

        console.log(`header.xml 파싱 완료: charPr=${this.charProperties.size}, paraPr=${this.paraProperties.size}, borderFill=${this.borderFills.size}`);
    }

    parseFontFaces(doc) {
        const fontfaceElems = doc.getElementsByTagNameNS(this.ns.hh, 'fontface');
        for (const ff of Array.from(fontfaceElems)) {
            const lang = ff.getAttribute('lang');
            const fonts = [];
            const fontElems = ff.getElementsByTagNameNS(this.ns.hh, 'font');
            for (const f of Array.from(fontElems)) {
                fonts.push({
                    id: f.getAttribute('id'),
                    face: f.getAttribute('face')
                });
            }
            this.fontFaces[lang] = fonts;
        }
    }

    parseBorderFills(doc) {
        const bfElems = doc.getElementsByTagNameNS(this.ns.hh, 'borderFill');
        for (const bf of Array.from(bfElems)) {
            const id = bf.getAttribute('id');
            const leftBorder = this.parseBorderElem(bf, 'leftBorder');
            const rightBorder = this.parseBorderElem(bf, 'rightBorder');
            const topBorder = this.parseBorderElem(bf, 'topBorder');
            const bottomBorder = this.parseBorderElem(bf, 'bottomBorder');

            // fillBrush에서 faceColor 추출
            let faceColor = '#FFFFFF';
            const winBrush = bf.getElementsByTagNameNS(this.ns.hc, 'winBrush')[0];
            if (winBrush) {
                faceColor = winBrush.getAttribute('faceColor') || '#FFFFFF';
            }

            this.borderFills.set(id, {
                id, leftBorder, rightBorder, topBorder, bottomBorder, faceColor
            });
        }
    }

    parseBorderElem(parent, name) {
        const elem = parent.getElementsByTagNameNS(this.ns.hh, name)[0];
        if (!elem) return { type: 'NONE', width: '0.1 mm', color: '#000000' };
        return {
            type: elem.getAttribute('type') || 'NONE',
            width: elem.getAttribute('width') || '0.1 mm',
            color: elem.getAttribute('color') || '#000000'
        };
    }

    parseCharProperties(doc) {
        const charPrElems = doc.getElementsByTagNameNS(this.ns.hh, 'charPr');
        for (const cp of Array.from(charPrElems)) {
            const id = cp.getAttribute('id');
            const height = parseInt(cp.getAttribute('height') || '1000');
            const textColor = cp.getAttribute('textColor') || '#000000';
            const shadeColor = cp.getAttribute('shadeColor') || 'none';
            const borderFillIDRef = cp.getAttribute('borderFillIDRef') || '1';

            // 폰트 참조
            const fontRef = cp.getElementsByTagNameNS(this.ns.hh, 'fontRef')[0];
            let hangulFontId = '0', latinFontId = '0';
            if (fontRef) {
                hangulFontId = fontRef.getAttribute('hangul') || '0';
                latinFontId = fontRef.getAttribute('latin') || '0';
            }

            // 굵게, 기울임 등
            const bold = cp.getElementsByTagNameNS(this.ns.hh, 'bold').length > 0;
            const italic = cp.getElementsByTagNameNS(this.ns.hh, 'italic').length > 0;
            const underline = cp.getElementsByTagNameNS(this.ns.hh, 'underline')[0] || null;
            const strikeout = cp.getElementsByTagNameNS(this.ns.hh, 'strikeout')[0] || null;
            const supscript = cp.getElementsByTagNameNS(this.ns.hh, 'supscript').length > 0;
            const subscript = cp.getElementsByTagNameNS(this.ns.hh, 'subscript').length > 0;

            // 폰트 이름 해석
            let hangulFont = '함초롬돋움', latinFont = '함초롬돋움';
            if (this.fontFaces['HANGUL'] && this.fontFaces['HANGUL'][parseInt(hangulFontId)]) {
                hangulFont = this.fontFaces['HANGUL'][parseInt(hangulFontId)].face;
            }
            if (this.fontFaces['LATIN'] && this.fontFaces['LATIN'][parseInt(latinFontId)]) {
                latinFont = this.fontFaces['LATIN'][parseInt(latinFontId)].face;
            }

            this.charProperties.set(id, {
                id, height, textColor, shadeColor, borderFillIDRef,
                bold, italic, underline, strikeout, supscript, subscript,
                hangulFont, latinFont
            });
        }
    }

    parseParaProperties(doc) {
        const paraPrElems = doc.getElementsByTagNameNS(this.ns.hh, 'paraPr');
        for (const pp of Array.from(paraPrElems)) {
            const id = pp.getAttribute('id');

            // 정렬
            const alignElem = pp.getElementsByTagNameNS(this.ns.hh, 'align')[0];
            let align = 'LEFT';
            if (alignElem) align = alignElem.getAttribute('horizontal') || 'LEFT';

            // 개요
            const headingElem = pp.getElementsByTagNameNS(this.ns.hh, 'heading')[0];
            let heading = 'NONE', level = '0', headingIdRef = '0';
            if (headingElem) {
                heading = headingElem.getAttribute('type') || 'NONE';
                level = headingElem.getAttribute('level') || '0';
                headingIdRef = headingElem.getAttribute('idRef') || '0';
            }

            // 여백
            let leftMargin = 0, rightMargin = 0, indent = 0, prevSpacing = 0, nextSpacing = 0;
            const marginElem = pp.getElementsByTagNameNS(this.ns.hh, 'margin')[0];
            if (marginElem) {
                const getVal = (ns, name) => {
                    const el = marginElem.getElementsByTagNameNS(ns, name)[0];
                    return el ? parseInt(el.getAttribute('value') || '0') : 0;
                };
                indent = getVal(this.ns.hc, 'intent');
                leftMargin = getVal(this.ns.hc, 'left');
                rightMargin = getVal(this.ns.hc, 'right');
                prevSpacing = getVal(this.ns.hc, 'prev');
                nextSpacing = getVal(this.ns.hc, 'next');
            }

            // 줄간격
            const lineSpacingElem = pp.getElementsByTagNameNS(this.ns.hh, 'lineSpacing')[0];
            let lineSpacingType = 'PERCENT', lineSpacingVal = 160;
            if (lineSpacingElem) {
                lineSpacingType = lineSpacingElem.getAttribute('type') || 'PERCENT';
                lineSpacingVal = parseInt(lineSpacingElem.getAttribute('value') || '160');
            }

            // 줄바꿈 설정
            const breakSetting = pp.getElementsByTagNameNS(this.ns.hh, 'breakSetting')[0];
            let keepWithNext = '0', keepLines = '0', pageBreakBefore = '0';
            if (breakSetting) {
                keepWithNext = breakSetting.getAttribute('keepWithNext') || '0';
                keepLines = breakSetting.getAttribute('keepLines') || '0';
                pageBreakBefore = breakSetting.getAttribute('pageBreakBefore') || '0';
            }

            // 테두리
            const borderElem = pp.getElementsByTagNameNS(this.ns.hh, 'border')[0];
            let borderFillIDRef = '1';
            if (borderElem) {
                borderFillIDRef = borderElem.getAttribute('borderFillIDRef') || '1';
            }

            this.paraProperties.set(id, {
                id, align, heading, level, headingIdRef,
                leftMargin, rightMargin, indent, prevSpacing, nextSpacing,
                lineSpacingType, lineSpacingVal,
                keepWithNext, keepLines, pageBreakBefore, borderFillIDRef
            });
        }
    }

    parseStyles(doc) {
        const styleElems = doc.getElementsByTagNameNS(this.ns.hh, 'style');
        for (const s of Array.from(styleElems)) {
            const id = s.getAttribute('id');
            const name = s.getAttribute('name') || '';
            const engName = s.getAttribute('engName') || '';
            const type = s.getAttribute('type') || 'PARA';
            this.styles.set(id, { id, name, engName, type });
        }
    }

    parseNumberings(doc) {
        const numElems = doc.getElementsByTagNameNS(this.ns.hh, 'numbering');
        for (const n of Array.from(numElems)) {
            const id = n.getAttribute('id');
            const levels = [];
            const paraHeadElems = n.getElementsByTagNameNS(this.ns.hh, 'paraHead');
            for (const ph of Array.from(paraHeadElems)) {
                levels.push({
                    level: ph.getAttribute('level') || '1',
                    numFormat: ph.getAttribute('numFormat') || 'DIGIT',
                    text: ph.getAttribute('text') || '',
                    start: ph.getAttribute('start') || '1'
                });
            }
            this.numberings.set(id, { id, levels });
        }
    }

    // ========================================
    // section0.xml 변환
    // ========================================
    convertSection(xml) {
        const doc = new DOMParser().parseFromString(xml, 'text/xml');
        const root = doc.documentElement;
        let bodyContent = '';

        // 첫 번째 문단에서 secPr 추출 (페이지 설정)
        const firstP = root.getElementsByTagNameNS(this.ns.hp, 'p')[0];
        if (firstP) {
            this.extractPageSetup(firstP);
        }

        // 모든 최상위 hp:p 요소 순회
        for (const child of Array.from(root.childNodes)) {
            if (child.nodeType !== 1) continue; // 엘리먼트 노드만
            if (child.localName === 'p') {
                bodyContent += this.convertParagraph(child);
            }
        }

        return bodyContent;
    }

    // 페이지 설정 추출
    extractPageSetup(firstP) {
        const secPr = firstP.getElementsByTagNameNS(this.ns.hp, 'secPr')[0];
        if (!secPr) return;

        const pagePr = secPr.getElementsByTagNameNS(this.ns.hp, 'pagePr')[0];
        if (pagePr) {
            const width = parseInt(pagePr.getAttribute('width') || '59528');
            const height = parseInt(pagePr.getAttribute('height') || '84188');
            // HWPUNIT → twips (÷5)
            this.pageWidth = Math.round(width / 5);
            this.pageHeight = Math.round(height / 5);

            const marginElem = pagePr.getElementsByTagNameNS(this.ns.hp, 'margin')[0];
            if (marginElem) {
                this.marginTop = Math.round(parseInt(marginElem.getAttribute('top') || '3545') / 5);
                this.marginBottom = Math.round(parseInt(marginElem.getAttribute('bottom') || '3545') / 5);
                this.marginLeft = Math.round(parseInt(marginElem.getAttribute('left') || '4250') / 5);
                this.marginRight = Math.round(parseInt(marginElem.getAttribute('right') || '4250') / 5);
                this.marginHeader = Math.round(parseInt(marginElem.getAttribute('header') || '2125') / 5);
                this.marginFooter = Math.round(parseInt(marginElem.getAttribute('footer') || '2125') / 5);
            }
        }
    }

    // ========================================
    // 문단 변환: hp:p → w:p
    // ========================================
    convertParagraph(pElem) {
        const paraPrIDRef = pElem.getAttribute('paraPrIDRef') || '0';
        const pageBreak = pElem.getAttribute('pageBreak') || '0';

        let xml = '<w:p>';

        // w:pPr 생성
        xml += this.getParaPropertiesXml(paraPrIDRef, pageBreak);

        // hp:run 순회
        const runs = pElem.getElementsByTagNameNS(this.ns.hp, 'run');
        for (const run of Array.from(runs)) {
            // run이 현재 p의 직접 자식인지 확인 (subList 내부 run 제외)
            if (run.parentNode !== pElem) continue;

            const charPrIDRef = run.getAttribute('charPrIDRef') || '0';

            // 표 처리
            const tables = run.getElementsByTagNameNS(this.ns.hp, 'tbl');
            if (tables.length > 0) {
                // 표가 있으면 현재 문단을 닫고 표를 삽입
                xml += '</w:p>';
                for (const tbl of Array.from(tables)) {
                    xml += this.convertTable(tbl);
                }
                xml += '<w:p>';
                xml += this.getParaPropertiesXml(paraPrIDRef, '0');
                continue;
            }

            // 이미지 처리
            const pics = run.getElementsByTagNameNS(this.ns.hp, 'pic');
            if (pics.length > 0) {
                for (const pic of Array.from(pics)) {
                    xml += this.convertImage(pic, charPrIDRef);
                }
                continue;
            }

            // 도형(rect) 내 텍스트 추출
            const rects = run.getElementsByTagNameNS(this.ns.hp, 'rect');
            if (rects.length > 0) {
                for (const rect of Array.from(rects)) {
                    xml += this.convertRect(rect, charPrIDRef);
                }
                continue;
            }

            // secPr, ctrl은 변환 대상이 아님 (skip)
            const secPr = run.getElementsByTagNameNS(this.ns.hp, 'secPr');
            if (secPr.length > 0) continue;

            // 일반 텍스트 처리
            const tElems = run.getElementsByTagNameNS(this.ns.hp, 't');
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
                    const tabs = tElem.getElementsByTagNameNS(this.ns.hp, 'tab');
                    for (const tab of Array.from(tabs)) {
                        xml += '<w:r>';
                        xml += this.getRunPropertiesXml(charPrIDRef);
                        xml += '<w:tab/>';
                        xml += '</w:r>';
                    }
                }
            } else {
                // 빈 run (텍스트 없음) → 빈 텍스트 run 생성
                // (secPr/ctrl이 아닌 경우에만)
                const ctrls = run.getElementsByTagNameNS(this.ns.hp, 'ctrl');
                if (ctrls.length === 0) {
                    xml += '<w:r>';
                    xml += this.getRunPropertiesXml(charPrIDRef);
                    xml += '</w:r>';
                }
            }
        }

        xml += '</w:p>';
        return xml;
    }

    // hp:t에서 순수 텍스트만 추출 (tab, lineseg 제외)
    getTextContent(tElem) {
        let text = '';
        for (const node of Array.from(tElem.childNodes)) {
            if (node.nodeType === 3) { // 텍스트 노드
                text += node.nodeValue;
            }
        }
        return text.trim();
    }

    // ========================================
    // w:pPr XML 생성
    // ========================================
    getParaPropertiesXml(paraPrId, pageBreak) {
        const pp = this.paraProperties.get(paraPrId);
        if (!pp && pageBreak !== '1') return '';

        let xml = '<w:pPr>';

        if (pp) {
            // 정렬
            const alignMap = { 'LEFT': 'left', 'CENTER': 'center', 'RIGHT': 'right', 'JUSTIFY': 'both', 'DISTRIBUTE': 'distribute' };
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

    // ========================================
    // w:rPr XML 생성
    // ========================================
    getRunPropertiesXml(charPrId) {
        const cp = this.charProperties.get(charPrId);
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
            const bf = this.borderFills.get(cp.borderFillIDRef);
            if (bf && bf.faceColor && bf.faceColor !== '#FFFFFF' && bf.faceColor !== 'none') {
                xml += `<w:shd w:val="clear" w:color="auto" w:fill="${bf.faceColor.replace('#', '')}"/>`;
            }
        }

        xml += '</w:rPr>';
        return xml;
    }

    // 색상 → highlight 이름 변환
    colorToHighlightName(color) {
        const map = {
            '#FFFF00': 'yellow', '#00FFFF': 'cyan', '#00FF00': 'green',
            '#FF00FF': 'magenta', '#0000FF': 'blue', '#FF0000': 'red',
            '#000000': 'black', '#FFFFFF': 'white', '#008000': 'darkGreen',
            '#808000': 'darkYellow', '#800000': 'darkRed', '#008080': 'darkCyan',
            '#800080': 'darkMagenta', '#00008B': 'darkBlue', '#A9A9A9': 'darkGray',
            '#D3D3D3': 'lightGray'
        };
        return map[color.toUpperCase()] || 'yellow';
    }

    // ========================================
    // 표 변환: hp:tbl → w:tbl
    // ========================================
    convertTable(tblElem) {
        let xml = '<w:tbl>';

        // 표 속성
        xml += '<w:tblPr>';
        xml += '<w:tblStyle w:val="TableGrid"/>';

        const szElem = tblElem.getElementsByTagNameNS(this.ns.hp, 'sz')[0];
        if (szElem) {
            const width = Math.round(parseInt(szElem.getAttribute('width') || '0') / 5);
            xml += `<w:tblW w:w="${width}" w:type="dxa"/>`;
        }

        xml += '<w:tblBorders>';
        xml += '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        xml += '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        xml += '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        xml += '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        xml += '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        xml += '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        xml += '</w:tblBorders>';
        xml += '<w:tblLook w:val="04A0"/>';
        xml += '</w:tblPr>';

        // 행 순회
        const rows = tblElem.getElementsByTagNameNS(this.ns.hp, 'tr');
        for (const row of Array.from(rows)) {
            xml += '<w:tr>';

            // 셀 순회
            const cells = row.getElementsByTagNameNS(this.ns.hp, 'tc');
            for (const cell of Array.from(cells)) {
                xml += this.convertTableCell(cell);
            }

            xml += '</w:tr>';
        }

        xml += '</w:tbl>';
        return xml;
    }

    convertTableCell(cellElem) {
        let xml = '<w:tc>';

        // 셀 속성
        xml += '<w:tcPr>';
        const cellSz = cellElem.getElementsByTagNameNS(this.ns.hp, 'cellSz')[0];
        if (cellSz) {
            const width = Math.round(parseInt(cellSz.getAttribute('width') || '0') / 5);
            xml += `<w:tcW w:w="${width}" w:type="dxa"/>`;
        }

        const colSpan = cellElem.getElementsByTagNameNS(this.ns.hp, 'cellSpan')[0];
        if (colSpan) {
            const cs = parseInt(colSpan.getAttribute('colSpan') || '1');
            if (cs > 1) xml += `<w:gridSpan w:val="${cs}"/>`;
            const rs = parseInt(colSpan.getAttribute('rowSpan') || '1');
            if (rs > 1) xml += '<w:vMerge w:val="restart"/>';
        }

        // 셀 배경색
        const bfIdRef = cellElem.getAttribute('borderFillIDRef');
        if (bfIdRef) {
            const bf = this.borderFills.get(bfIdRef);
            if (bf && bf.faceColor && bf.faceColor !== '#FFFFFF' && bf.faceColor !== 'none') {
                xml += `<w:shd w:val="clear" w:color="auto" w:fill="${bf.faceColor.replace('#', '')}"/>`;
            }
        }

        xml += '</w:tcPr>';

        // 셀 내 문단 변환 (subList 내의 hp:p)
        const subList = cellElem.getElementsByTagNameNS(this.ns.hp, 'subList')[0];
        if (subList) {
            const paragraphs = subList.getElementsByTagNameNS(this.ns.hp, 'p');
            let hasParagraph = false;
            for (const p of Array.from(paragraphs)) {
                // subList의 직접 자식 p만 변환
                if (p.parentNode === subList) {
                    xml += this.convertParagraph(p);
                    hasParagraph = true;
                }
            }
            if (!hasParagraph) {
                xml += '<w:p/>';
            }
        } else {
            xml += '<w:p/>';
        }

        xml += '</w:tc>';
        return xml;
    }

    // ========================================
    // 이미지 변환: hp:pic → w:drawing
    // ========================================
    convertImage(picElem, charPrId) {
        const imgElem = picElem.getElementsByTagNameNS(this.ns.hc, 'img')[0];
        if (!imgElem) return '';

        const binaryItemIDRef = imgElem.getAttribute('binaryItemIDRef');
        if (!binaryItemIDRef || !this.images.has(binaryItemIDRef)) return '';

        const imgInfo = this.images.get(binaryItemIDRef);

        // 이미지 크기 (HWPUNIT → EMU: × 914400 / 7200 = × 127)
        const szElem = picElem.getElementsByTagNameNS(this.ns.hp, 'sz')[0];
        let cx = 5000000, cy = 3000000; // 기본값
        if (szElem) {
            const w = parseInt(szElem.getAttribute('width') || '0');
            const h = parseInt(szElem.getAttribute('height') || '0');
            if (w > 0) cx = w * 127;
            if (h > 0) cy = h * 127;
        }

        // 이미지 relationship 등록
        const relId = `rId${++this.relIdCounter}`;
        const fileName = `image${this.imageRels.length + 1}.${imgInfo.ext}`;
        this.imageRels.push({
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
        xml += '<wp:docPr id="1" name="Picture"/>';
        xml += '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">';
        xml += '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">';
        xml += '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">';
        xml += '<pic:nvPicPr><pic:cNvPr id="0" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>';
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

    // ========================================
    // 도형(rect) 변환: 텍스트만 추출
    // ========================================
    convertRect(rectElem, charPrId) {
        let xml = '';
        const drawText = rectElem.getElementsByTagNameNS(this.ns.hp, 'drawText')[0];
        if (!drawText) return '';

        const subList = drawText.getElementsByTagNameNS(this.ns.hp, 'subList')[0];
        if (!subList) return '';

        // subList 내의 hp:p에서 텍스트 추출
        const paragraphs = subList.getElementsByTagNameNS(this.ns.hp, 'p');
        for (const p of Array.from(paragraphs)) {
            if (p.parentNode !== subList) continue;
            const runs = p.getElementsByTagNameNS(this.ns.hp, 'run');
            for (const run of Array.from(runs)) {
                if (run.parentNode !== p) continue;
                const cId = run.getAttribute('charPrIDRef') || charPrId;
                const tElems = run.getElementsByTagNameNS(this.ns.hp, 't');
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

    // ========================================
    // DOCX 패키지 생성
    // ========================================
    async createDocxPackage(bodyXml) {
        const zip = new JSZip();

        // [Content_Types].xml
        zip.file('[Content_Types].xml', this.createContentTypes());

        // _rels/.rels
        zip.file('_rels/.rels', this.createTopRels());

        // word/document.xml
        zip.file('word/document.xml', this.createDocumentXml(bodyXml));

        // word/_rels/document.xml.rels
        zip.file('word/_rels/document.xml.rels', this.createDocumentRels());

        // word/styles.xml
        zip.file('word/styles.xml', this.createStylesXml());

        // word/settings.xml
        zip.file('word/settings.xml', this.createSettingsXml());

        // word/fontTable.xml
        zip.file('word/fontTable.xml', this.createFontTableXml());

        // 이미지 파일 추가
        for (const rel of this.imageRels) {
            zip.file(`word/${rel.target}`, rel.data);
        }

        return await zip.generateAsync({ type: 'nodebuffer' });
    }

    createContentTypes() {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        xml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        xml += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        xml += '<Default Extension="xml" ContentType="application/xml"/>';
        xml += '<Default Extension="png" ContentType="image/png"/>';
        xml += '<Default Extension="jpg" ContentType="image/jpeg"/>';
        xml += '<Default Extension="jpeg" ContentType="image/jpeg"/>';
        xml += '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>';
        xml += '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>';
        xml += '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>';
        xml += '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>';
        xml += '</Types>';
        return xml;
    }

    createTopRels() {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        xml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>';
        xml += '</Relationships>';
        return xml;
    }

    createDocumentXml(bodyXml) {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        xml += '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"';
        xml += ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
        xml += ' xmlns:o="urn:schemas-microsoft-com:office:office"';
        xml += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
        xml += ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
        xml += ' xmlns:v="urn:schemas-microsoft-com:vml"';
        xml += ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
        xml += ' xmlns:w10="urn:schemas-microsoft-com:office:word"';
        xml += ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
        xml += ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
        xml += ' xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"';
        xml += ' xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"';
        xml += ' xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"';
        xml += ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">';
        xml += '<w:body>';
        xml += bodyXml;

        // sectPr (페이지 설정)
        xml += '<w:sectPr>';
        xml += `<w:pgSz w:w="${this.pageWidth}" w:h="${this.pageHeight}"/>`;
        xml += `<w:pgMar w:top="${this.marginTop}" w:right="${this.marginRight}" w:bottom="${this.marginBottom}" w:left="${this.marginLeft}" w:header="${this.marginHeader}" w:footer="${this.marginFooter}" w:gutter="0"/>`;
        xml += '</w:sectPr>';

        xml += '</w:body>';
        xml += '</w:document>';
        return xml;
    }

    createDocumentRels() {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        xml += '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        xml += '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>';
        xml += '<Relationship Id="rIdFontTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>';

        // 이미지 relationship
        for (const rel of this.imageRels) {
            xml += `<Relationship Id="${rel.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${rel.target}"/>`;
        }

        xml += '</Relationships>';
        return xml;
    }

    createStylesXml() {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        xml += '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';

        // 기본 스타일
        xml += '<w:docDefaults>';
        xml += '<w:rPrDefault><w:rPr>';
        xml += '<w:rFonts w:ascii="맑은 고딕" w:eastAsia="맑은 고딕" w:hAnsi="맑은 고딕"/>';
        xml += '<w:sz w:val="20"/><w:szCs w:val="20"/>';
        xml += '</w:rPr></w:rPrDefault>';
        xml += '<w:pPrDefault><w:pPr>';
        xml += '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>';
        xml += '</w:pPr></w:pPrDefault>';
        xml += '</w:docDefaults>';

        // Normal 스타일
        xml += '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">';
        xml += '<w:name w:val="Normal"/>';
        xml += '</w:style>';

        // TableGrid 스타일 (표용)
        xml += '<w:style w:type="table" w:styleId="TableGrid">';
        xml += '<w:name w:val="Table Grid"/>';
        xml += '<w:tblPr><w:tblBorders>';
        xml += '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        xml += '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        xml += '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        xml += '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        xml += '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        xml += '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
        xml += '</w:tblBorders></w:tblPr>';
        xml += '</w:style>';

        xml += '</w:styles>';
        return xml;
    }

    createSettingsXml() {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '<w:defaultTabStop w:val="720"/>' +
            '<w:compat><w:useFELayout/></w:compat>' +
            '</w:settings>';
    }

    createFontTableXml() {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        xml += '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';

        // 사용된 폰트 수집
        const usedFonts = new Set(['맑은 고딕']);
        for (const [, cp] of this.charProperties) {
            if (cp.hangulFont) usedFonts.add(cp.hangulFont);
            if (cp.latinFont) usedFonts.add(cp.latinFont);
        }

        for (const font of usedFonts) {
            xml += `<w:font w:name="${this.escapeXml(font)}">`;
            xml += '<w:charset w:val="81"/>';
            xml += '<w:family w:val="modern"/>';
            xml += '<w:pitch w:val="variable"/>';
            xml += '</w:font>';
        }

        xml += '</w:fonts>';
        return xml;
    }

    // ========================================
    // 유틸리티
    // ========================================
    escapeXml(text) {
        if (!text) return '';
        return text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }
}

// ========================================
// CLI 실행
// ========================================
async function main() {
    const args = process.argv.slice(2);
    if (args.length < 1) {
        console.log('사용법: node hwpx-to-docx.js <hwpx파일_또는_디렉토리> [출력파일.docx]');
        console.log('예시: node hwpx-to-docx.js ~/바탕화면/input/ output.docx');
        process.exit(1);
    }

    const input = args[0];
    const output = args[1] || 'output.docx';

    try {
        const converter = new HwpxToDocxConverter();
        const docxBuffer = await converter.convert(input);
        fs.writeFileSync(output, docxBuffer);
        console.log(`✅ 변환 완료: ${output}`);
        console.log(`   파일 크기: ${(docxBuffer.length / 1024).toFixed(1)} KB`);
    } catch (err) {
        console.error('❌ 변환 오류:', err.message);
        console.error(err.stack);
        process.exit(1);
    }
}

main();

module.exports = HwpxToDocxConverter;
