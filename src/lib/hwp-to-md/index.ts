import { HwpParser } from './parser';

export async function convertHwpToMd(file: File | ArrayBuffer | Blob): Promise<string> {
    let arrayBuffer: ArrayBuffer;

    if (file instanceof File || file instanceof Blob) {
        arrayBuffer = await file.arrayBuffer();
    } else {
        arrayBuffer = file;
    }

    const data = new Uint8Array(arrayBuffer);

    try {
        const parser = new HwpParser(data);
        return parser.parse();
    } catch (err: any) {
        console.error("HWP to MD conversion error parsing file", err);
        throw new Error(`Failed to convert HWP: ${err.message}`);
    }
}
