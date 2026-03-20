import { useState } from 'react';
import { convertHwpxToMd } from '../lib/hwpx-to-md';

export const useHwpxToMd = () => {
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState<Error | null>(null);
    const [result, setResult] = useState<string | null>(null);

    const convert = async (file: File | ArrayBuffer | Blob) => {
        setIsLoading(true);
        setError(null);
        setResult(null);
        try {
            const markdown = await convertHwpxToMd(file);
            setResult(markdown);
            return markdown;
        } catch (err: any) {
            setError(err);
            throw err;
        } finally {
            setIsLoading(false);
        }
    };

    return { convert, isLoading, error, result };
};
