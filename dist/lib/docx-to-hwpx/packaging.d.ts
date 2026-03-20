import { ConversionState } from './state';

export declare class HwpxPackager {
    private state;
    constructor(state: ConversionState);
    packageHwpx(hwpxData: any): Promise<Blob>;
}
