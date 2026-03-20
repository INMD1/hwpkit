import JSZip from 'jszip';
import { HwpParser } from './parser';
import { DocxGenerator } from './generator';
import type { ImageRel } from './parser';

/**
 * HWP → DOCX 변환 오케스트레이터
 * HwpParser로 HWP 바이너리를 파싱하고,
 * DocxGenerator로 DOCX XML을 생성한 뒤,
 * 최종 DOCX 패키지(ZIP)를 만듭니다.
 */
export class HwpToDocxConverter {
    async convert(input: File | ArrayBuffer | Blob): Promise<Blob> {
        let arrayBuffer: ArrayBuffer;
        if (input instanceof File || input instanceof Blob) {
            arrayBuffer = await input.arrayBuffer();
        } else {
            arrayBuffer = input;
        }

        const data = new Uint8Array(arrayBuffer);

        // 1. HWP 바이너리 파싱
        const parser = new HwpParser();
        const parseResult = parser.parse(data);

        // 2. DOCX XML 생성기 초기화
        const generator = new DocxGenerator(
            parseResult.faceNames,
            parseResult.charShapes,
            parseResult.paraShapes,
            parseResult.pageDefs,
            parseResult.borderFills,
        );

        // 3. 섹션별 본문 XML 생성 (머리말/꼬리말 포함)
        const sectionStreams = parser.getSectionStreams();
        let bodyXml = '';
        let headerXml: string | null = null;
        let footerXml: string | null = null;
        for (const stream of sectionStreams) {
            const result = generator.parseSection(stream);
            bodyXml += result.body;
            if (result.header !== null) headerXml = result.header;
            if (result.footer !== null) footerXml = result.footer;
        }
        bodyXml += generator.buildSectPrXml(headerXml !== null, footerXml !== null);

        // 4. DOCX 패키지 생성 (문서에서 파싱된 폰트 목록 전달)
        return this.createDocxPackage(bodyXml, parseResult.imageRels, headerXml, footerXml, parseResult.faceNames);
    }

    // ─── DOCX 패키징 ──────────────────────────────────────────

    private async createDocxPackage(
        bodyXml: string,
        imageRels: ImageRel[],
        headerXml: string | null,
        footerXml: string | null,
        faceNames: string[] = [],
    ): Promise<Blob> {
        const zip = new JSZip();
        const xmlDecl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
            ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

        zip.file('[Content_Types].xml', this.contentTypesXml(imageRels, headerXml !== null, footerXml !== null));
        zip.file('_rels/.rels', this.topRelsXml());
        zip.file('word/document.xml', this.documentXml(bodyXml));
        zip.file('word/_rels/document.xml.rels', this.documentRelsXml(imageRels, headerXml !== null, footerXml !== null));
        zip.file('word/styles.xml', this.stylesXml(faceNames));
        zip.file('word/settings.xml',
            `${xmlDecl}<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
            '<w:zoom w:percent="100"/></w:settings>'
        );
        zip.file('word/fontTable.xml', this.fontTableXml(faceNames));

        // 머리말 파일
        if (headerXml !== null) {
            zip.file('word/header1.xml', `${xmlDecl}<w:hdr ${NS}>${headerXml}</w:hdr>`);
        }
        // 꼬리말 파일
        if (footerXml !== null) {
            zip.file('word/footer1.xml', `${xmlDecl}<w:ftr ${NS}>${footerXml}</w:ftr>`);
        }

        // 이미지 파일 추가
        for (const rel of imageRels) {
            zip.file(`word/${rel.target}`, rel.data);
        }

        return await zip.generateAsync({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
    }

    // ─── MIME 타입 매핑 ───────────────────────────────────────
    private static readonly MIME_MAP: Record<string, string> = {
        png: 'image/png',
        jpg: 'image/jpeg',
        jpeg: 'image/jpeg',
        bmp: 'image/bmp',
        gif: 'image/gif',
        tiff: 'image/tiff',
        tif: 'image/tiff',
        emf: 'image/x-emf',
        wmf: 'image/x-wmf',
        svg: 'image/svg+xml',
        webp: 'image/webp',
    };

    // ─── 한국어 폰트 판별 ─────────────────────────────────────
    private isKoreanFont(name: string): boolean {
        return /함초롬|맑은|나눔|굴림|바탕|돋움|궁서|HCR|새굴림|윤|한컴|[가-힣]/.test(name);
    }

    // ─── 기본 폰트 결정 (문서 faceNames 기반) ─────────────────
    private getDefaultFonts(faceNames: string[]): { korean: string; latin: string } {
        const valid = faceNames.filter(f => f && f.trim());
        const korean = valid.find(f => this.isKoreanFont(f)) || 'Malgun Gothic';
        const latin  = valid.find(f => !this.isKoreanFont(f) && /^[A-Za-z]/.test(f)) || korean;
        return { korean, latin };
    }

    private contentTypesXml(imageRels: ImageRel[], hasHeader = false, hasFooter = false): string {
        // 실제 사용된 이미지 확장자만 동적 등록
        const usedExts = new Set<string>();
        for (const rel of imageRels) {
            if (rel.ext) usedExts.add(rel.ext.toLowerCase());
        }

        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>';

        for (const ext of usedExts) {
            const mime = HwpToDocxConverter.MIME_MAP[ext] ?? `image/${ext}`;
            xml += `<Default Extension="${ext}" ContentType="${mime}"/>`;
        }

        xml += '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
            '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
            '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>';
        if (hasHeader) xml += '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';
        if (hasFooter) xml += '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>';
        xml += '</Types>';
        return xml;
    }

    private topRelsXml(): string {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
            '</Relationships>';
    }

    private documentXml(bodyXml: string): string {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
            ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            `<w:body>${bodyXml}</w:body></w:document>`;
    }

    private documentRelsXml(imageRels: ImageRel[], hasHeader = false, hasFooter = false): string {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        xml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        xml += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>';
        xml += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>';
        if (hasHeader) xml += '<Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>';
        if (hasFooter) xml += '<Relationship Id="rIdFtr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>';
        for (const rel of imageRels) {
            xml += `<Relationship Id="${rel.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${rel.target}"/>`;
        }
        xml += '</Relationships>';
        return xml;
    }

    /** styles.xml: 문서의 실제 폰트를 기본값으로 사용 */
    private stylesXml(faceNames: string[] = []): string {
        const { korean, latin } = this.getDefaultFonts(faceNames);
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '<w:docDefaults><w:rPrDefault><w:rPr>' +
            `<w:rFonts w:ascii="${latin}" w:eastAsia="${korean}" w:hAnsi="${latin}" w:cs="${korean}"/>` +
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults></w:styles>';
    }

    /** fontTable.xml: 문서에 사용된 폰트 전체 목록 등록 */
    private fontTableXml(faceNames: string[] = []): string {
        const xmlDecl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        const uniqueFonts = [...new Set(faceNames.filter(f => f && f.trim()))];
        if (uniqueFonts.length === 0) uniqueFonts.push('Malgun Gothic');

        let xml = `${xmlDecl}<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`;
        for (const font of uniqueFonts) {
            // 한국어 폰트: charset 0x81(한글), 그 외: 0x00
            const charset = this.isKoreanFont(font) ? '81' : '00';
            xml += `<w:font w:name="${font}">` +
                `<w:charset w:val="${charset}"/>` +
                `<w:family w:val="auto"/>` +
                `<w:pitch w:val="variable"/>` +
                `</w:font>`;
        }
        xml += '</w:fonts>';
        return xml;
    }
}
