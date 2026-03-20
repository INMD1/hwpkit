import { HwpxToDocxConverter } from './converter';

export const convertHwpxToDocx = async (file: File | ArrayBuffer | Blob): Promise<Blob> => {
    const converter = new HwpxToDocxConverter();
    return await converter.convert(file);
};
