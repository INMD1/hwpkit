import { useState } from 'react';
import { convertHwpxToDocx } from '../lib/hwpx-to-docx';

export const useHwpxToDocx = () => {
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState<Error | null>(null);
    const [result, setResult] = useState<Blob | null>(null);

    const convert = async (file: File | ArrayBuffer | Blob) => {
        setIsLoading(true);
        setError(null);
        setResult(null);
        try {
            const blob = await convertHwpxToDocx(file);
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
