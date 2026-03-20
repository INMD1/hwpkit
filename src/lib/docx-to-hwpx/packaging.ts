import JSZip from 'jszip';
import { ConversionState } from './state';

export class HwpxPackager {
    private state: ConversionState;

    constructor(state: ConversionState) {
        this.state = state;
    }

    public async packageHwpx(hwpxData: any): Promise<Blob> {
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
        for (const [, imgInfo] of this.state.images) {
            zip.file(imgInfo.path, imgInfo.data);
        }

        console.log(`✅ Packaged HWPX with ${this.state.images.size} images`);

        return await zip.generateAsync({ type: 'blob' });
    }
}
