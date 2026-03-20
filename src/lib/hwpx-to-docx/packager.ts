import JSZip from 'jszip';
import { ConversionContext } from './context';

export class DocxPackager {
    private ctx: ConversionContext;

    constructor(ctx: ConversionContext) {
        this.ctx = ctx;
    }

    public async createDocxPackage(bodyXml: string): Promise<Blob> {
        const zip = new JSZip();

        zip.file('[Content_Types].xml', this.createContentTypes());
        zip.file('_rels/.rels', this.createTopRels());
        zip.file('word/document.xml', this.createDocumentXml(bodyXml));
        zip.file('word/_rels/document.xml.rels', this.createDocumentRels());
        zip.file('word/styles.xml', this.createStylesXml());
        zip.file('word/settings.xml', this.createSettingsXml());
        zip.file('word/fontTable.xml', this.createFontTableXml());

        for (const rel of this.ctx.imageRels) {
            zip.file(`word/${rel.target}`, rel.data);
        }

        return await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    }

    // ─── MIME 타입 매핑 ──────────────────────────────────────
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

    // ─── 한국어 폰트 판별 ────────────────────────────────────
    private isKoreanFont(name: string): boolean {
        return /함초롬|맑은|나눔|굴림|바탕|돋움|궁서|HCR|새굴림|윤|한컴|[가-힣]/.test(name);
    }

    // ─── 문서에서 사용된 폰트 목록 수집 ────────────────────────
    private getDocumentFonts(): string[] {
        const fonts = new Set<string>();
        for (const cp of this.ctx.charProperties.values()) {
            if (cp.hangulFont) fonts.add(cp.hangulFont);
            if (cp.latinFont)  fonts.add(cp.latinFont);
        }
        // fontFaces에서도 수집
        for (const faces of Object.values(this.ctx.fontFaces)) {
            for (const f of faces) {
                if (f.face) fonts.add(f.face);
            }
        }
        return [...fonts].filter(f => f.trim());
    }

    // ─── 기본 폰트 결정 ─────────────────────────────────────
    private getDefaultFonts(): { korean: string; latin: string } {
        const fonts = this.getDocumentFonts();
        const korean = fonts.find(f => this.isKoreanFont(f)) ?? 'Malgun Gothic';
        const latin  = fonts.find(f => !this.isKoreanFont(f) && /^[A-Za-z]/.test(f)) ?? korean;
        return { korean, latin };
    }

    private createContentTypes(): string {
        // 실제 사용된 이미지 확장자만 동적 등록
        const usedExts = new Set<string>();
        for (const rel of this.ctx.imageRels) {
            if (rel.ext) usedExts.add(rel.ext.toLowerCase());
        }

        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        xml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        xml += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        xml += '<Default Extension="xml" ContentType="application/xml"/>';
        for (const ext of usedExts) {
            const mime = DocxPackager.MIME_MAP[ext] ?? `image/${ext}`;
            xml += `<Default Extension="${ext}" ContentType="${mime}"/>`;
        }
        xml += '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>';
        xml += '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>';
        xml += '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>';
        xml += '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>';
        xml += '</Types>';
        return xml;
    }

    private createTopRels(): string {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
    }

    private createDocumentXml(bodyXml: string): string {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${bodyXml}<w:sectPr><w:pgSz w:w="${this.ctx.pageWidth}" w:h="${this.ctx.pageHeight}"/><w:pgMar w:top="${this.ctx.marginTop}" w:right="${this.ctx.marginRight}" w:bottom="${this.ctx.marginBottom}" w:left="${this.ctx.marginLeft}" w:header="${this.ctx.marginHeader}" w:footer="${this.ctx.marginFooter}" w:gutter="0"/></w:sectPr></w:body></w:document>`;
    }

    private createDocumentRels(): string {
        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        xml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        xml += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>';
        xml += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>';

        for (const rel of this.ctx.imageRels) {
            xml += `<Relationship Id="${rel.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${rel.target.replace('word/', '')}"/>`;
        }

        xml += '</Relationships>';
        return xml;
    }

    /** styles.xml: 문서의 실제 폰트를 기본값으로 사용 */
    private createStylesXml(): string {
        const { korean, latin } = this.getDefaultFonts();
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            '<w:docDefaults><w:rPrDefault><w:rPr>' +
            `<w:rFonts w:ascii="${latin}" w:eastAsia="${korean}" w:hAnsi="${latin}" w:cs="${korean}"/>` +
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults></w:styles>';
    }

    private createSettingsXml(): string {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/></w:settings>';
    }

    /** fontTable.xml: 문서에 사용된 폰트 전체 목록 등록 */
    private createFontTableXml(): string {
        const fonts = this.getDocumentFonts();
        if (fonts.length === 0) fonts.push('Malgun Gothic');

        let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
        xml += '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
        for (const font of fonts) {
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
