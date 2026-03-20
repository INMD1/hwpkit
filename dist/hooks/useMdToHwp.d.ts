export declare const useMdToHwp: () => {
    convert: (file: File | string | Blob) => Promise<Blob | null>;
    isConverting: boolean;
    error: string | null;
};
