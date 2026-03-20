export function escapeXml(text: string): string {
    if (!text) return '';
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

export function getFontFamilyType(face: string): string {
    if (isMonospaceFont(face)) return 'FCAT_FIXED';
    if (face.includes('고딕') || face.includes('Gothic') || face.includes('Sans') || face.includes('Arial') ||
        face.includes('Helvetica') || face.includes('Verdana') || face.includes('Tahoma') || face.includes('돋움')) return 'FCAT_GOTHIC';
    const serifFonts = ['바탕', '명조', 'Batang', 'Times', 'Serif', 'Georgia'];
    if (serifFonts.some(s => face.includes(s))) return 'FCAT_MYEONGJO';
    return 'FCAT_UNKNOWN';
}

export function isMonospaceFont(face: string): boolean {
    const monoFonts = ['Courier New', 'Consolas', 'Lucida Console', 'Monaco', 'Menlo',
        'monospace', 'Ubuntu Mono', 'Source Code Pro', 'Fira Code', 'JetBrains Mono', 'D2Coding', 'NanumGothicCoding'];
    return monoFonts.some(m => face.toLowerCase().includes(m.toLowerCase()));
}

export function getImageDimensions(data: Uint8Array, ext: string) {
    let width = 0;
    let height = 0;

    try {
        if (ext === 'png') {
            const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
            if (view.getUint32(0) === 0x89504E47 && view.getUint32(4) === 0x0D0A1A0A) {
                width = view.getUint32(16);
                height = view.getUint32(20);
            }
        } else if (ext === 'jpg' || ext === 'jpeg') {
            const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
            let offset = 2; // Skip SOI (FF D8)
            while (offset < view.byteLength) {
                if (offset + 1 >= view.byteLength) break;

                const marker = view.getUint16(offset);
                offset += 2;

                if (marker === 0xFFC0 || marker === 0xFFC2) { // SOF0 or SOF2
                    height = view.getUint16(offset + 3);
                    width = view.getUint16(offset + 5);
                    break;
                } else if ((marker & 0xFF00) === 0xFF00 && marker !== 0xFFFF) {
                    const len = view.getUint16(offset);
                    offset += len;
                } else {
                    break; // Invalid marker or padding
                }
            }
        }
    } catch (e) {
        console.warn('이미지 크기 추출 실패 (기본값 사용):', e);
    }

    // Fallback or default
    return { width: width || 100, height: height || 100 };
}
