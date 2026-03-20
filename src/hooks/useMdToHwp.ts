import { useState, useCallback } from 'react';
import { convertMdToHwp } from '../lib/md-to-hwp';

export const useMdToHwp = () => {
    const [isConverting, setIsConverting] = useState(false);
    const [error, setError] = useState<string | null>(null);

    const convert = useCallback(async (file: File | string | Blob): Promise<Blob | null> => {
        setIsConverting(true);
        setError(null);

        try {
            const blob = await convertMdToHwp(file);
            return blob;
        } catch (err: any) {
            console.error('MD to HWP Conversion failed:', err);
            setError(err.message || 'Conversion failed');
            return null;
        } finally {
            setIsConverting(false);
        }
    }, []);

    return { convert, isConverting, error };
};
