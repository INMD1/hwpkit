import { HwpToDocxConverter } from './converter';

export { HwpToDocxConverter };

export async function convertHwpToDocx(file: File | ArrayBuffer | Blob): Promise<Blob> {
    const converter = new HwpToDocxConverter();
    return converter.convert(file);
}
