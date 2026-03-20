import { HwpxToMdConverter } from './converter';

export async function convertHwpxToMd(file: File | ArrayBuffer | Blob): Promise<string> {
    try {
        const converter = new HwpxToMdConverter();
        return await converter.convert(file);
    } catch (err: any) {
        console.error('HWPX to MD conversion error:', err);
        throw new Error(`Failed to convert HWPX: ${err.message}`);
    }
}
