import JSZip from 'jszip';
import { ConversionContext } from './context';
import { HwpxParser } from './parser';
import { DocxGenerator } from './generator';
import { DocxPackager } from './packager';

export class HwpxToDocxConverter {
    async convert(hwpxInput: File | ArrayBuffer | Blob): Promise<Blob> {
        let zip: JSZip;
        if (hwpxInput instanceof File || hwpxInput instanceof Blob) {
            zip = await JSZip.loadAsync(hwpxInput);
        } else {
            zip = await JSZip.loadAsync(hwpxInput);
        }

        return await this.convertFromZip(zip);
    }

    private async convertFromZip(zip: JSZip): Promise<Blob> {
        console.log('HWPX ZIP에서 변환...');

        const ctx = new ConversionContext();
        const parser = new HwpxParser(ctx);
        const generator = new DocxGenerator(ctx);
        const packager = new DocxPackager(ctx);

        const contentHpf = zip.file('Contents/content.hpf');
        if (contentHpf) {
            await parser.parseContentHpf(await contentHpf.async('string'));
        }

        const header = zip.file('Contents/header.xml');
        if (header) {
            parser.parseHeader(await header.async('string'));
        }

        const binDataFolder = zip.folder('BinData');
        if (binDataFolder) {
            const imageFiles: { name: string, file: JSZip.JSZipObject }[] = [];
            binDataFolder.forEach((relativePath, file) => {
                if (!file.dir) imageFiles.push({ name: relativePath, file });
            });
            for (const { name, file } of imageFiles) {
                const data = await file.async('uint8array');
                const parts = name.split('.');
                const ext = parts.length > 1 ? parts.pop()!.toLowerCase() : '';
                const id = parts.length > 0 ? parts[0].split('/').pop()! : name;

                const mediaType = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`;
                ctx.images.set(id, { data, mediaType, ext });
            }
        }

        // 모든 섹션 파일을 번호 순서대로 처리 (section0.xml, section1.xml, ...)
        const sectionPaths = Object.keys(zip.files)
            .filter(k => /Contents\/section\d+\.xml$/i.test(k))
            .sort((a, b) => {
                const na = parseInt(a.match(/section(\d+)/i)?.[1] || '0', 10);
                const nb = parseInt(b.match(/section(\d+)/i)?.[1] || '0', 10);
                return na - nb;
            });

        let bodyXml = '';
        for (const path of sectionPaths) {
            const secFile = zip.file(path);
            if (secFile) {
                bodyXml += generator.convertSection(await secFile.async('string'));
            }
        }

        return await packager.createDocxPackage(bodyXml);
    }
}
