export declare const useHwpxToDocx: () => {
    convert: (file: File | ArrayBuffer | Blob) => Promise<Blob>;
    isLoading: boolean;
    error: Error | null;
    result: Blob | null;
};
