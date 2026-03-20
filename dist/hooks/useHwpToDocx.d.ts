export declare const useHwpToDocx: () => {
    convert: (file: File | ArrayBuffer | Blob) => Promise<Blob>;
    isLoading: boolean;
    error: Error | null;
    result: Blob | null;
};
