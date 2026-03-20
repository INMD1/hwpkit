import { DocxToHwpxConverter } from './converter';

export const convertDocxToHwpx = async (file: File | ArrayBuffer | Blob): Promise<Blob> => {
    const converter = new DocxToHwpxConverter();
    return await converter.convert(file);
};
