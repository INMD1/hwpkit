import { useState } from 'react';
import { convertDocxToHwpx } from '../lib/docx-to-hwpx';

export const useDocxToHwpx = () => {
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState<Error | null>(null);
    const [result, setResult] = useState<Blob | null>(null);

    const convert = async (file: File | ArrayBuffer | Blob) => {
        setIsLoading(true);
        setError(null);
        setResult(null);
        try {
            const blob = await convertDocxToHwpx(file);
            setResult(blob);
            return blob;
        } catch (err: any) {
            setError(err);
            throw err;
        } finally {
            setIsLoading(false);
        }
    };

    return { convert, isLoading, error, result };
};
