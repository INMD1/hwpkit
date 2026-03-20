export declare const useHwpToMd: () => {
    convert: (file: File | ArrayBuffer | Blob) => Promise<string>;
    isLoading: boolean;
    error: Error | null;
    result: string | null;
};
