import JSZip from 'jszip';
import { ConversionState } from './state';
import { DocxParser } from './parser';
import { HwpxGenerator } from './generator';
import { HwpxPackager } from './packaging';

export class DocxToHwpxConverter {
    async convert(docxInput: File | ArrayBuffer | Blob): Promise<Blob> {
        console.log('[1/8] DOCX 로드 중...');
        const zip = await JSZip.loadAsync(docxInput);

        const state = new ConversionState();
        const parser = new DocxParser(state);
        const generator = new HwpxGenerator(state);
        const packager = new HwpxPackager(state);

        console.log('[2/8] 메타데이터 파싱...');
        await parser.parseMetadata(zip);

        console.log('[3/8] 이미지 및 관계 파싱...');
        await parser.parseRelationships(zip);
        await parser.parseImages(zip);
        await parser.parseStyles(zip);
        await parser.parseNumbering(zip);
        await parser.parseFootnotes(zip);
        await parser.parseEndnotes(zip);

        console.log('[4/8] DOCX 파싱 중...');
        const docXml = await zip.file('word/document.xml')?.async('text');
        if (!docXml) throw new Error('word/document.xml not found');

        const domParser = new DOMParser();
        const doc = domParser.parseFromString(docXml, 'text/xml');

        // Page Setup Parsing
        parser.parsePageSetup(doc);

        console.log('[5/8] 스타일 초기화...');
        // Styles are already initialized in ConversionState constructor

        console.log('[6/8] HWPX 생성 중...');
        const hwpxData = generator.createHwpx(doc);

        console.log('[7/8] 패키징 중...');
        const blob = await packager.packageHwpx(hwpxData);

        console.log('[8/8] 완료!');
        return blob;
    }
}
