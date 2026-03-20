export declare const useDocxToHwpx: () => {
    convert: (file: File | ArrayBuffer | Blob) => Promise<Blob>;
    isLoading: boolean;
    error: Error | null;
    result: Blob | null;
};
